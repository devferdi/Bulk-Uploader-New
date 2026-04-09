# Shopify Agency Setup Guide

This repo supports the client-store OAuth flow:

1. A partner-owned Shopify app with `Custom distribution`
2. A small OAuth backend that stores one offline token per client store
3. The desktop bulk tool fetching the token from that backend

## 1. Create the Shopify app

1. Log in to your Partner Dev Dashboard.
2. Create a new app for the bulk tool.
3. Choose `Custom distribution` so the app stays off the App Store.
4. Set the app URL to your backend root, for example:
   - `https://shopify-bulk-oauth.youragency.com/`
5. Add the redirect URL:
   - `https://shopify-bulk-oauth.youragency.com/auth/callback`
6. Configure the access scopes your desktop app needs.

Recommended starting scopes for this repo:

- `read_products`
- `write_products`
- `read_inventory`
- `write_inventory`
- `read_locations`
- `read_files`
- `write_files`

Add more only if testing shows they are required for your specific metafield or collection workflows.

## 2. Run the OAuth backend

1. Copy [backend/.env.example](/Users/ferdinand/Documents/Upwork/Hux%20Agency/Biotta/Bulk%20Uploader%20Biotta/backend/.env.example) to `.env` in the `backend` folder.
2. Fill in the Shopify client ID, Shopify client secret, public app URL, and agency API key.
3. Create a virtual environment and install dependencies:

```bash
cd backend
python3 -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

4. Start the backend locally:

```bash
uvicorn app:app --reload
```

5. For real Shopify installs, run it on a public HTTPS URL or through an HTTPS tunnel.
6. Deploy the same backend to a public HTTPS URL for production.

## 3. Connect a client store

1. Generate a custom install link in the Shopify Partner Dashboard for the client store.
2. Have the merchant install the app, or install it yourself if they granted you the right permissions.
3. After install, either:
   - click `Open app` from Shopify admin, or
   - open `https://your-backend-domain/auth/start?shop=client-store.myshopify.com`
4. Complete the Shopify permission screen.
5. The backend stores the offline token in its SQLite database.

## 4. Point the desktop app at the backend

1. Create `credentials.txt` in the repo root.
2. Fill in:
   - `store_name`
   - `oauth_backend_url`
   - `agency_api_key`
   - `openai_api_key` if you use the SEO alt text feature
3. Remove `shopify_client_id`, `shopify_client_secret`, and `access_token` for stores that use the backend flow.
4. Launch the desktop tool normally.

The desktop app now tries this order:

1. Use `access_token` from `credentials.txt` if present
2. Otherwise fetch the token from `/api/shops/{shop}/token` on the OAuth backend
3. Otherwise try the same-organization Dev Dashboard client-credentials flow

## 5. Recommended next steps

1. Host the backend behind HTTPS with a persistent SQLite file or move token storage to Postgres.
2. Register the `app/uninstalled` webhook to:
   - `https://your-backend-domain/webhooks/app/uninstalled`
3. Move the desktop app off old Shopify API versions store by store.
4. Later, migrate more REST calls to GraphQL.
