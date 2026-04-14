from __future__ import annotations

import threading
import traceback
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from alt_text_service import (
    APP_NAME,
    DEFAULT_LANGUAGE,
    DEFAULT_MAX_WORDS,
    DEFAULT_MODEL,
    AltTextGenerationError,
    GeneratorSettings,
    generate_alt_texts,
    load_initial_settings,
    save_settings,
)


APP_DIR = Path(__file__).resolve().parent
REPO_ROOT = APP_DIR.parent.parent


class AltTextGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("920x760")
        self.root.minsize(820, 680)

        initial_settings = load_initial_settings(REPO_ROOT)

        self.input_path_var = tk.StringVar()
        self.output_dir_var = tk.StringVar()
        self.brand_name_var = tk.StringVar(value=initial_settings.brand_name)
        self.api_key_var = tk.StringVar(value=initial_settings.api_key)
        self.model_var = tk.StringVar(value=initial_settings.model or DEFAULT_MODEL)
        self.language_var = tk.StringVar(
            value=initial_settings.language or DEFAULT_LANGUAGE,
        )
        self.max_words_var = tk.StringVar(
            value=str(initial_settings.max_words or DEFAULT_MAX_WORDS),
        )
        self.overwrite_var = tk.BooleanVar(
            value=initial_settings.overwrite_existing_new_alt_text,
        )
        self.status_var = tk.StringVar(
            value="Choose a spreadsheet from Files Alt Text download to begin.",
        )

        self.worker_thread: threading.Thread | None = None
        self.is_running = False

        self._build_layout()

    def _build_layout(self) -> None:
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(1, weight=1)

        header = ttk.Frame(self.root, padding=(18, 18, 18, 10))
        header.grid(row=0, column=0, sticky="ew")
        header.columnconfigure(0, weight=1)

        title = ttk.Label(
            header,
            text=APP_NAME,
            font=("SF Pro Display", 22, "bold"),
        )
        title.grid(row=0, column=0, sticky="w")

        subtitle = ttk.Label(
            header,
            text=(
                "Generate OpenAI alt text from the Shopify files workbook and save a new XLSX with New Alt Text filled in."
            ),
            wraplength=720,
            foreground="#5f6368",
        )
        subtitle.grid(row=1, column=0, sticky="w", pady=(6, 0))

        content = ttk.Frame(self.root, padding=(18, 0, 18, 18))
        content.grid(row=1, column=0, sticky="nsew")
        content.columnconfigure(0, weight=1)
        content.rowconfigure(1, weight=1)

        settings_frame = ttk.LabelFrame(content, text="Inputs", padding=16)
        settings_frame.grid(row=0, column=0, sticky="ew")
        settings_frame.columnconfigure(1, weight=1)

        self._add_path_row(
            settings_frame,
            row=0,
            label="Spreadsheet",
            variable=self.input_path_var,
            button_label="Browse",
            command=self.choose_input_file,
        )
        self._add_path_row(
            settings_frame,
            row=1,
            label="Output folder",
            variable=self.output_dir_var,
            button_label="Choose",
            command=self.choose_output_folder,
        )
        self._add_entry_row(
            settings_frame,
            row=2,
            label="Brand name",
            variable=self.brand_name_var,
        )
        self._add_entry_row(
            settings_frame,
            row=3,
            label="OpenAI API key",
            variable=self.api_key_var,
            show="*",
        )
        self._add_entry_row(
            settings_frame,
            row=4,
            label="Model",
            variable=self.model_var,
        )
        self._add_entry_row(
            settings_frame,
            row=5,
            label="Language",
            variable=self.language_var,
        )
        self._add_entry_row(
            settings_frame,
            row=6,
            label="Max words",
            variable=self.max_words_var,
        )

        overwrite_checkbox = ttk.Checkbutton(
            settings_frame,
            text="Overwrite existing values in the New Alt Text column",
            variable=self.overwrite_var,
        )
        overwrite_checkbox.grid(row=7, column=0, columnspan=3, sticky="w", pady=(10, 0))

        action_row = ttk.Frame(settings_frame)
        action_row.grid(row=8, column=0, columnspan=3, sticky="ew", pady=(14, 0))
        action_row.columnconfigure(3, weight=1)

        self.save_button = ttk.Button(
            action_row,
            text="Save settings",
            command=self.handle_save_settings,
        )
        self.save_button.grid(row=0, column=0, sticky="w")

        self.start_button = ttk.Button(
            action_row,
            text="Generate alt texts",
            command=self.start_generation,
        )
        self.start_button.grid(row=0, column=1, sticky="w", padx=(10, 0))

        clear_button = ttk.Button(
            action_row,
            text="Clear log",
            command=self.clear_log,
        )
        clear_button.grid(row=0, column=2, sticky="w", padx=(10, 0))

        status_label = ttk.Label(
            action_row,
            textvariable=self.status_var,
            foreground="#5f6368",
        )
        status_label.grid(row=0, column=3, sticky="e", padx=(16, 0))

        log_frame = ttk.LabelFrame(content, text="Progress", padding=16)
        log_frame.grid(row=1, column=0, sticky="nsew", pady=(16, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_output = scrolledtext.ScrolledText(
            log_frame,
            wrap=tk.WORD,
            font=("SF Mono", 12),
            height=20,
        )
        self.log_output.grid(row=0, column=0, sticky="nsew")

    def _add_path_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        button_label: str,
        command,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(0, 10))
        entry = ttk.Entry(parent, textvariable=variable)
        entry.grid(row=row, column=1, sticky="ew", padx=(12, 10), pady=(0, 10))
        button = ttk.Button(parent, text=button_label, command=command)
        button.grid(row=row, column=2, sticky="ew", pady=(0, 10))

    def _add_entry_row(
        self,
        parent: ttk.LabelFrame,
        row: int,
        label: str,
        variable: tk.StringVar,
        show: str | None = None,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", pady=(0, 10))
        entry = ttk.Entry(parent, textvariable=variable, show=show or "")
        entry.grid(row=row, column=1, columnspan=2, sticky="ew", padx=(12, 0), pady=(0, 10))

    def log(self, message: str) -> None:
        def append() -> None:
            self.log_output.insert(tk.END, f"{message}\n")
            self.log_output.see(tk.END)

        self.root.after(0, append)

    def clear_log(self) -> None:
        self.log_output.delete("1.0", tk.END)

    def choose_input_file(self) -> None:
        path = filedialog.askopenfilename(
            title="Choose the Shopify files alt-text spreadsheet",
            filetypes=[("Excel files", "*.xlsx")],
        )
        if not path:
            return
        self.input_path_var.set(path)
        if not self.output_dir_var.get():
            self.output_dir_var.set(str(Path(path).resolve().parent))

    def choose_output_folder(self) -> None:
        path = filedialog.askdirectory(title="Choose output folder")
        if path:
            self.output_dir_var.set(path)

    def build_settings(self) -> GeneratorSettings:
        max_words_text = self.max_words_var.get().strip() or str(DEFAULT_MAX_WORDS)
        try:
            max_words = int(max_words_text)
        except ValueError as exc:
            raise AltTextGenerationError("Max words must be a whole number.") from exc

        return GeneratorSettings(
            api_key=self.api_key_var.get().strip(),
            brand_name=self.brand_name_var.get().strip(),
            language=self.language_var.get().strip() or DEFAULT_LANGUAGE,
            max_words=max_words,
            model=self.model_var.get().strip() or DEFAULT_MODEL,
            overwrite_existing_new_alt_text=self.overwrite_var.get(),
        )

    def handle_save_settings(self) -> None:
        try:
            settings = self.build_settings()
            settings_path = save_settings(settings)
        except Exception as exc:
            messagebox.showerror("Save failed", str(exc))
            return

        self.status_var.set("Settings saved locally on this Mac.")
        self.log(f"Saved settings to {settings_path}")

    def set_running(self, value: bool) -> None:
        self.is_running = value
        state = tk.DISABLED if value else tk.NORMAL
        self.start_button.config(state=state)
        self.save_button.config(state=state)

    def start_generation(self) -> None:
        if self.is_running:
            return

        spreadsheet_path = Path(self.input_path_var.get().strip())
        output_dir_text = self.output_dir_var.get().strip()
        output_dir = Path(output_dir_text) if output_dir_text else spreadsheet_path.parent

        try:
            settings = self.build_settings()
        except Exception as exc:
            messagebox.showerror("Invalid settings", str(exc))
            return

        self.set_running(True)
        self.status_var.set("Generating alt texts...")
        self.log("")
        self.log("Starting OpenAI alt-text generation...")

        def worker() -> None:
            try:
                save_settings(settings)
                result_path = generate_alt_texts(
                    spreadsheet_path=spreadsheet_path,
                    output_dir=output_dir,
                    settings=settings,
                    log=self.log,
                )
            except Exception as exc:
                traceback.print_exc()
                error = exc
                self.root.after(0, lambda error=error: self.finish_with_error(error))
                return

            self.root.after(0, lambda: self.finish_successfully(result_path))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def finish_successfully(self, result_path: Path) -> None:
        self.set_running(False)
        self.status_var.set("Done")
        self.log(f"Finished successfully: {result_path}")
        messagebox.showinfo(
            "Alt text generation complete",
            f"Saved the updated workbook here:\n\n{result_path}",
        )

    def finish_with_error(self, error: Exception) -> None:
        self.set_running(False)
        self.status_var.set("Failed")
        self.log(f"Generation failed: {error}")
        messagebox.showerror("Generation failed", str(error))


def main() -> None:
    root = tk.Tk()
    ttk.Style(root)
    app = AltTextGeneratorApp(root)
    app.log("Ready.")
    root.mainloop()


if __name__ == "__main__":
    main()
