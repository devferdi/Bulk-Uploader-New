# Deploy HUX Bulk Loader To A Real Store

This app is now prepared to be deployed as a single hosted service that includes:

- the Shopify embedded admin app
- the Python spreadsheet worker
- the existing `Shopify Bulk Tool.py` logic

## 1. Host the app

Use a host that can deploy a Dockerfile from the repo root.

Recommended Dockerfile:

```text
/Users/ferdinand/Documents/Upwork/Hux Agency/Biotta/Bulk Uploader Biotta/Dockerfile.shopify-admin-app
```

Fastest option for this repo:

- use the root-level `render.yaml`
- create a Render Blueprint from the Git repo
- let it deploy the `hux-bulk-loader` Docker service and attach the disk at `/var/data`

The app now exposes a dedicated health endpoint at `/healthz` for hosted health checks.

## 2. Required environment variables

If you deploy without the Blueprint, set these on the host:

```text
SHOPIFY_API_KEY=
SHOPIFY_API_SECRET=
SCOPES=read_files,write_files,write_inventory,read_inventory,read_locations,read_metaobject_definitions,read_metaobjects,write_metaobjects,read_products,write_products
SHOPIFY_APP_URL=https://your-real-hostname.example
DATABASE_URL=file:./prisma/dev.sqlite
NODE_ENV=production
```

For Render with the Blueprint:

- `SHOPIFY_API_KEY` and `SHOPIFY_API_SECRET` are prompted as secrets
- `SCOPES`, `NODE_ENV`, `PORT`, and `DATABASE_URL` are prefilled
- `SHOPIFY_APP_URL` is optional at runtime because the app falls back to Render's `RENDER_EXTERNAL_URL`

For a real production setup, replace the SQLite `DATABASE_URL` with a managed database URL if your host does not provide persistent disk storage.

## 3. Update Shopify app config

Edit:

```text
shopify.app.hux-bulk-loader.toml
```

Set:

- `application_url = "https://your-real-hostname.example"`
- `redirect_urls = ["https://your-real-hostname.example/auth/callback"]`

## 4. Deploy app config to Shopify

From:

```text
/Users/ferdinand/Documents/Upwork/Hux Agency/Biotta/Bulk Uploader Biotta/shopify-admin-app/biotta-bulk-admin
```

Run:

```bash
shopify app deploy --config shopify.app.hux-bulk-loader.toml
```

## 5. Install on the real store

In the Shopify Partner / Dev Dashboard:

1. Open the app.
2. Set distribution to `Custom distribution`.
3. Add `biotta-ag.myshopify.com`.
4. Generate the install link.
5. Open the link as the store owner and install the app.
