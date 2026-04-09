# Shopify Dev Dashboard Setup Guide

This desktop tool can keep working without legacy custom apps by using a Shopify app created in the Dev Dashboard.

## 1. Create the app in Shopify Dev Dashboard

1. Open the store organization's Shopify Dev Dashboard.
2. Create a new app for this bulk uploader.
3. Add the Admin API scopes this tool needs, then release a version.
4. Install the app on the target store from the app's Home tab.
5. Copy the app's client ID and client secret from the app settings.

Recommended starting scopes for this repo:

- `read_products`
- `write_products`
- `read_inventory`
- `write_inventory`
- `read_locations`
- `read_files`
- `write_files`

Add collection or metafield scopes if testing shows they are needed for your store workflows.

## 2. Configure the desktop app

1. Copy [credentials.example.txt](/Users/ferdinand/Documents/Upwork/Hux%20Agency/Biotta/Bulk%20Uploader%20Biotta/credentials.example.txt) to `credentials.txt`.
2. Fill in:
   - `store_name`
   - `shopify_client_id`
   - `shopify_client_secret`
   - `openai_api_key` if you use the SEO alt text feature
3. Remove `access_token` once the store is migrated to the Dev Dashboard flow.

## 3. Run the desktop app

The app now authenticates like this:

1. Use `access_token` from `credentials.txt` if present
2. Otherwise request a fresh Admin API token from Shopify using the Dev Dashboard app credentials

Shopify's Dev Dashboard access tokens expire after about 24 hours, so the desktop app requests a fresh token whenever it starts a Shopify operation.

## 4. Notes

- This flow is for stores owned by the same organization as the Dev Dashboard app.
- If the app is not installed, the client secret is wrong, or the scopes were not released, Shopify will reject the token request.
- The code still uses a lot of REST Admin API calls, so the default API version in this repo is `2026-01`.
