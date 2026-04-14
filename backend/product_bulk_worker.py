import argparse
import importlib.util
import json
import sys
from pathlib import Path


REPO_ROOT = Path(__file__).resolve().parents[1]
BULK_TOOL_PATH = REPO_ROOT / "Shopify Bulk Tool.py"


def load_bulk_tool_module():
    spec = importlib.util.spec_from_file_location("shopify_bulk_tool", BULK_TOOL_PATH)
    if spec is None or spec.loader is None:
        raise RuntimeError(f"Could not load bulk tool module from {BULK_TOOL_PATH}")

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def build_context(module, shop, access_token, api_version):
    return module.build_shopify_context_from_values(
        shop_name=shop,
        access_token=access_token,
        api_version=api_version,
    )


def emit_result(action, file_path):
    print(
        "WORKER_RESULT_JSON="
        + json.dumps(
            {
                "action": action,
                "file_path": file_path,
            }
        )
    )


def parse_args():
    parser = argparse.ArgumentParser(
        description="Run product spreadsheet download/upload logic without the desktop GUI."
    )
    parser.add_argument(
        "action",
        choices=[
            "download",
            "upload",
            "products-download",
            "products-upload",
            "metaobjects-download",
            "metaobjects-upload",
            "collections-download",
            "collections-upload",
            "file-alt-texts-download",
            "file-alt-texts-upload",
        ],
    )
    parser.add_argument("--shop", required=True, help="Shop domain or shop name.")
    parser.add_argument("--access-token", required=True, help="Admin API access token.")
    parser.add_argument(
        "--api-version",
        default="2026-01",
        help="Shopify Admin API version to use.",
    )
    parser.add_argument(
        "--file",
        help="Path to the spreadsheet to upload when action=upload.",
    )
    parser.add_argument(
        "--output-dir",
        help="Directory where downloaded spreadsheets should be written.",
    )
    parser.add_argument(
        "--script-dir",
        default=str(REPO_ROOT),
        help="Repo root used to resolve local assets and credentials fallback.",
    )
    return parser.parse_args()


def main():
    args = parse_args()
    module = load_bulk_tool_module()
    shopify_context = build_context(
        module,
        shop=args.shop,
        access_token=args.access_token,
        api_version=args.api_version,
    )

    if args.action in {"download", "products-download"}:
        output_path = module.run_downloader_logic(
            shopify_context=shopify_context,
            script_dir=args.script_dir,
            output_dir=args.output_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action in {"upload", "products-upload"}:
        if not args.file:
            raise RuntimeError("--file is required when action=upload")

        output_path = module.run_uploader_logic(
            file_path=args.file,
            shopify_context=shopify_context,
            script_dir=args.script_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action == "metaobjects-download":
        output_path = module.metaobject_run_downloader_logic(
            shopify_context=shopify_context,
            script_dir=args.script_dir,
            output_dir=args.output_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action == "metaobjects-upload":
        if not args.file:
            raise RuntimeError("--file is required when action=metaobjects-upload")

        output_path = module.metaobject_run_uploader_logic(
            file_path=args.file,
            shopify_context=shopify_context,
            script_dir=args.script_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action == "collections-download":
        output_path = module.collection_run_downloader_logic(
            shopify_context=shopify_context,
            script_dir=args.script_dir,
            output_dir=args.output_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action == "collections-upload":
        if not args.file:
            raise RuntimeError("--file is required when action=collections-upload")

        output_path = module.collection_run_uploader_logic(
            file_path=args.file,
            shopify_context=shopify_context,
            script_dir=args.script_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action == "file-alt-texts-download":
        output_path = module.download_shopify_files_alt_texts(
            shopify_context=shopify_context,
            script_dir=args.script_dir,
            output_dir=args.output_dir,
        )
        emit_result(args.action, output_path)
        return 0

    if args.action == "file-alt-texts-upload":
        if not args.file:
            raise RuntimeError("--file is required when action=file-alt-texts-upload")

        output_path = module.upload_shopify_files_alt_texts(
            file_path=args.file,
            shopify_context=shopify_context,
            script_dir=args.script_dir,
        )
        emit_result(args.action, output_path)
        return 0

    raise RuntimeError(f"Unsupported bulk worker action: {args.action}")


if __name__ == "__main__":
    try:
        raise SystemExit(main())
    except Exception as exc:
        print(
            "WORKER_RESULT_JSON="
            + json.dumps(
                {
                    "error": str(exc),
                }
            )
        )
        raise
