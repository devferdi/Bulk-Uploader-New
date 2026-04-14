# HUX Alt Text Generator

Standalone Mac desktop tool for generating `New Alt Text` values with the OpenAI API from the Shopify files workbook.

## What it does

1. Open the workbook exported from the `Files Alt Text` flow.
2. Read the `Filename`, `Alt Text`, and `URL` columns.
3. Use OpenAI to generate improved values for `New Alt Text`.
4. Save a new `.xlsx` file without touching Shopify directly.

## Run it on your Mac

From this repo:

```bash
python3 desktop_apps/hux_alt_text_generator/app.py
```

Or double-click:

```text
desktop_apps/hux_alt_text_generator/HUX Alt Text Generator.command
```

## Inputs

- `Spreadsheet`: the Shopify files alt-text export workbook
- `Output folder`: where the generated workbook should be saved
- `Brand name`: optional brand/store name to weave into the alt text when natural
- `OpenAI API key`: used to call the OpenAI API
- `Model`: defaults to `gpt-4o`
- `Language`: defaults to `German`
- `Max words`: defaults to `15`

## Saved settings

The app stores its settings here on macOS:

```text
~/Library/Application Support/HUX Alt Text Generator/settings.json
```

## Spreadsheet format

The input workbook must contain these columns:

- `Filename`
- `Alt Text`
- `URL`

If `New Alt Text` is missing, the app will add it automatically.

## Notes

- This app is independent from Shopify auth.
- It only works on the spreadsheet you already exported.
- It writes periodic backup files while generating alt text.
