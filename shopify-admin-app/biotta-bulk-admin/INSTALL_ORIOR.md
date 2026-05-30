# Install ORIOR As A Separate Shopify App

Use this flow when `orior-ag` is not in the same Shopify Plus organization as
`biotta-ag`.

## Why this needs a separate app

This codebase can serve multiple clients, but each deployed app instance only
supports one Shopify API key and secret pair at runtime:

- `SHOPIFY_API_KEY`
- `SHOPIFY_API_SECRET`

That means ORIOR needs:

1. its own Shopify app record in the Partner Dashboard
2. its own hosted deployment instance
3. its own `shopify.app.orior.toml` config

## Setup

### 1. Create the ORIOR app in Shopify Partner Dashboard

Create a new app:

- app name: `ORIOR Bulk Loader`
- distribution: `Custom distribution`

After the app exists, copy its client ID.

### 2. Create a second hosted instance

Duplicate the existing Render service or create a new web service from the same
repo using:

- Dockerfile: `Dockerfile.shopify-admin-app`
- a new hostname, for example `https://orior-bulk-loader.onrender.com`

Set these environment variables on the ORIOR service:

- `SHOPIFY_API_KEY=<ORIOR app client ID>`
- `SHOPIFY_API_SECRET=<ORIOR app client secret>`
- `SCOPES=read_files,write_files,read_content,write_content,write_inventory,read_inventory,read_locations,read_metaobject_definitions,read_metaobjects,write_metaobjects,read_products,write_products`
- `DATABASE_URL=file:/var/data/dev.sqlite`
- `NODE_ENV=production`
- `PORT=3000`

### 3. Fill in the ORIOR Shopify config

Edit `shopify.app.orior.toml` and replace:

- `REPLACE_WITH_ORIOR_CLIENT_ID`
- `https://REPLACE_WITH_ORIOR_HOSTNAME`

Optional shortcut:

```bash
npm run config:link:orior -- --client-id=<ORIOR_CLIENT_ID>
```

That pulls the app record into `shopify.app.orior.toml`, after which you only
need to make sure `application_url` and `redirect_urls` match the ORIOR host.

Example:

```toml
client_id = "your-orior-client-id"
application_url = "https://orior-bulk-loader.onrender.com"

[auth]
redirect_urls = [ "https://orior-bulk-loader.onrender.com/auth/callback" ]
```

### 4. Deploy the ORIOR app config to Shopify

From this directory:

```text
/Users/ferdinand/Documents/Upwork/Hux Agency/Biotta/Bulk Uploader Biotta/shopify-admin-app/biotta-bulk-admin
```

Run:

```bash
npm run deploy:orior
```

### 5. Generate the install link

In the ORIOR app record in Shopify Partner Dashboard:

1. Open `Distribution`
2. Keep `Custom distribution`
3. Generate the install link
4. Open the link as the owner/admin of `orior-ag`

## Troubleshooting

If Shopify shows `invalid_link_organization`, the install link belongs to the
wrong app record or wrong merchant organization. Generate a fresh link from the
new `ORIOR Bulk Loader` app, not from the Biotta app.
