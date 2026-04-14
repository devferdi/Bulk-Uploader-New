from __future__ import annotations

import json
import threading
import time
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Callable

import pandas as pd
from openai import OpenAI


APP_NAME = "HUX Alt Text Generator"
DEFAULT_LANGUAGE = "German"
DEFAULT_MAX_WORDS = 15
DEFAULT_MODEL = "gpt-4o"
DEFAULT_TEMPERATURE = 0.2
BACKUP_INTERVAL = 5
REQUIRED_COLUMNS = ("Filename", "Alt Text", "URL")
SETTINGS_DIR = Path.home() / "Library" / "Application Support" / APP_NAME
SETTINGS_PATH = SETTINGS_DIR / "settings.json"
file_lock = threading.Lock()

LogCallback = Callable[[str], None]
CancelCallback = Callable[[], bool]


class AltTextGenerationError(RuntimeError):
    pass


@dataclass
class GeneratorSettings:
    api_key: str = ""
    brand_name: str = ""
    language: str = DEFAULT_LANGUAGE
    max_words: int = DEFAULT_MAX_WORDS
    model: str = DEFAULT_MODEL
    overwrite_existing_new_alt_text: bool = False
    temperature: float = DEFAULT_TEMPERATURE


def set_dataframe_cell(df: pd.DataFrame, row_index: int, column: str, value: str) -> None:
    try:
        df.at[row_index, column] = value
    except (TypeError, ValueError, pd.errors.LossySetitemError):
        df[column] = df[column].astype(object)
        df.at[row_index, column] = value


def read_key_value_file(file_path: Path) -> dict[str, str]:
    values: dict[str, str] = {}
    if not file_path.exists():
        return values

    for raw_line in file_path.read_text(encoding="utf-8").splitlines():
        line = raw_line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, value = line.split("=", 1)
        values[key.strip()] = value.strip()
    return values


def prettify_brand_name(value: str) -> str:
    cleaned = " ".join((value or "").replace("-", " ").replace("_", " ").split())
    return cleaned.title()


def discover_repo_credentials(base_dir: Path) -> dict[str, str]:
    candidates = [
        base_dir / "credentials.txt",
        base_dir.parent / "credentials.txt",
        base_dir.parent.parent / "credentials.txt",
    ]

    for candidate in candidates:
        values = read_key_value_file(candidate)
        if values:
            return values

    return {}


def load_saved_settings() -> GeneratorSettings:
    if not SETTINGS_PATH.exists():
        return GeneratorSettings()

    try:
        payload = json.loads(SETTINGS_PATH.read_text(encoding="utf-8"))
    except (json.JSONDecodeError, OSError):
        return GeneratorSettings()

    known_fields = {field.name for field in GeneratorSettings.__dataclass_fields__.values()}
    filtered_payload = {
        key: value
        for key, value in payload.items()
        if key in known_fields
    }
    return GeneratorSettings(**filtered_payload)


def load_initial_settings(base_dir: Path) -> GeneratorSettings:
    saved_settings = load_saved_settings()
    credentials = discover_repo_credentials(base_dir)

    api_key = saved_settings.api_key or credentials.get("openai_api_key", "")
    brand_name = saved_settings.brand_name or prettify_brand_name(
        credentials.get("store_name", ""),
    )

    return GeneratorSettings(
        api_key=api_key,
        brand_name=brand_name,
        language=saved_settings.language or DEFAULT_LANGUAGE,
        max_words=saved_settings.max_words or DEFAULT_MAX_WORDS,
        model=saved_settings.model or DEFAULT_MODEL,
        overwrite_existing_new_alt_text=saved_settings.overwrite_existing_new_alt_text,
        temperature=saved_settings.temperature or DEFAULT_TEMPERATURE,
    )


def save_settings(settings: GeneratorSettings) -> Path:
    SETTINGS_DIR.mkdir(parents=True, exist_ok=True)
    SETTINGS_PATH.write_text(
        json.dumps(asdict(settings), indent=2),
        encoding="utf-8",
    )
    return SETTINGS_PATH


def validate_inputs(
    spreadsheet_path: Path,
    output_dir: Path,
    settings: GeneratorSettings,
) -> None:
    if not spreadsheet_path.exists():
        raise AltTextGenerationError("Choose an existing XLSX spreadsheet first.")

    if spreadsheet_path.suffix.lower() != ".xlsx":
        raise AltTextGenerationError("The input file must be an .xlsx spreadsheet.")

    if not settings.api_key.strip():
        raise AltTextGenerationError("Add your OpenAI API key before starting.")

    if not settings.model.strip():
        raise AltTextGenerationError("Add an OpenAI model name before starting.")

    if not settings.language.strip():
        raise AltTextGenerationError("Choose an output language before starting.")

    if settings.max_words <= 0:
        raise AltTextGenerationError("Max words must be greater than zero.")

    output_dir.mkdir(parents=True, exist_ok=True)


def build_messages(
    filename: str,
    old_alt_text: str,
    image_url: str,
    settings: GeneratorSettings,
) -> list[dict]:
    brand_instruction = ""
    if settings.brand_name.strip():
        brand_instruction = (
            f"Integrate the brand name '{settings.brand_name.strip()}' naturally when it truly fits the image. "
        )

    system_prompt = (
        "You are an SEO-focused alt text writer. "
        f"Write concise alt text in {settings.language.strip()}. "
        f"Keep it to a maximum of {settings.max_words} words. "
        "Describe only what is visibly present in the image. "
        "Use natural language, not keyword stuffing. "
        f"{brand_instruction}"
        "Avoid hallucinating product details that are not visible. "
        "Return only the alt text."
    )

    user_text = (
        "Create one improved alt text for this image.\n"
        f"Filename: {filename or 'Unknown'}\n"
        f"Existing alt text: {old_alt_text or 'None'}\n"
        f"Language: {settings.language.strip()}\n"
        f"Max words: {settings.max_words}\n"
        "Describe visible objects, materials, colors, context, and activity only when clearly present."
    )

    return [
        {
            "role": "system",
            "content": system_prompt,
        },
        {
            "role": "user",
            "content": [
                {
                    "type": "text",
                    "text": user_text,
                },
                {
                    "type": "image_url",
                    "image_url": {
                        "url": image_url,
                    },
                },
            ],
        },
    ]


def extract_response_text(response) -> str:
    try:
        content = response.choices[0].message.content
    except (AttributeError, IndexError, KeyError, TypeError) as exc:
        raise AltTextGenerationError(
            "OpenAI returned an unexpected response while generating alt text."
        ) from exc

    if isinstance(content, str):
        result = content.strip()
    elif isinstance(content, list):
        text_parts = []
        for item in content:
            if isinstance(item, dict) and item.get("type") == "text":
                text_parts.append(str(item.get("text", "")))
        result = " ".join(text_parts).strip()
    else:
        result = str(content).strip()

    if not result:
        raise AltTextGenerationError("OpenAI returned an empty alt text.")

    return result


def generate_alt_texts(
    spreadsheet_path: Path,
    output_dir: Path,
    settings: GeneratorSettings,
    log: LogCallback,
    should_cancel: CancelCallback | None = None,
) -> Path:
    validate_inputs(spreadsheet_path, output_dir, settings)

    df = pd.read_excel(spreadsheet_path)
    missing_columns = [column for column in REQUIRED_COLUMNS if column not in df.columns]
    if missing_columns:
        raise AltTextGenerationError(
            "Spreadsheet is missing required columns: "
            + ", ".join(missing_columns)
        )

    if "New Alt Text" not in df.columns:
        df["New Alt Text"] = ""

    client = OpenAI(api_key=settings.api_key.strip())
    current_time = datetime.now().strftime("%Y%m%d_%H%M%S")
    temp_file = output_dir / f"temp_alt_texts_{current_time}.xlsx"
    final_file = output_dir / f"shopify_images_with_new_alt_texts_{current_time}.xlsx"

    total_rows = len(df.index)
    log(f"Processing {total_rows} rows from {spreadsheet_path.name}...")

    for row_index, row in df.iterrows():
        if should_cancel and should_cancel():
            raise AltTextGenerationError("Generation was cancelled.")

        image_url = row.get("URL")
        filename = str(row.get("Filename") or "").strip()
        existing_alt_text = str(row.get("Alt Text") or "").strip()
        current_new_alt_text = str(row.get("New Alt Text") or "").strip()

        if not image_url or not isinstance(image_url, str):
            log(f"Skipping row {row_index + 2}: missing image URL.")
            continue

        if current_new_alt_text and not settings.overwrite_existing_new_alt_text:
            log(f"Skipping row {row_index + 2}: New Alt Text already exists.")
            continue

        try:
            response = client.chat.completions.create(
                model=settings.model.strip(),
                messages=build_messages(
                    filename=filename,
                    old_alt_text=existing_alt_text,
                    image_url=image_url,
                    settings=settings,
                ),
                temperature=settings.temperature,
                max_tokens=120,
            )
            generated_text = extract_response_text(response)
            set_dataframe_cell(df, row_index, "New Alt Text", generated_text)
            log(f"Generated alt text for row {row_index + 2}: {generated_text}")

            if row_index % BACKUP_INTERVAL == 0:
                with file_lock:
                    df.to_excel(temp_file, index=False)
                log(f"Saved backup at row {row_index + 2}.")
        except Exception as exc:
            log(f"Error on row {row_index + 2}: {exc}")

    for attempt in range(3):
        try:
            with file_lock:
                df.to_excel(final_file, index=False)
            log(f"Saved final workbook to {final_file}")
            return final_file
        except PermissionError:
            if attempt == 2:
                break
            log(
                f"Final save was blocked by macOS permissions. Retrying ({attempt + 2}/3)..."
            )
            time.sleep(2)
        except Exception as exc:
            raise AltTextGenerationError(
                f"Could not save the final workbook: {exc}"
            ) from exc

    raise AltTextGenerationError(
        f"Could not save the final workbook after 3 attempts. Backup file: {temp_file}"
    )
