import base64
import hashlib
import hmac
import os
import secrets
import sqlite3
import urllib.parse
from datetime import datetime, timezone

import requests
from fastapi import FastAPI, HTTPException, Request
from fastapi.responses import HTMLResponse, JSONResponse, RedirectResponse


DEFAULT_SHOPIFY_API_VERSION = "2026-01"
STATE_COOKIE_NAME = "shopify_oauth_state"


app = FastAPI(title="Shopify Bulk Tool OAuth Backend")


def utcnow_iso():
    return datetime.now(timezone.utc).isoformat()


def get_env(name, default=None, required=False):
    value = os.getenv(name, default)
    if required and not value:
        raise RuntimeError(f"Missing required environment variable: {name}")
    return value


def get_client_id():
    return get_env("SHOPIFY_CLIENT_ID", required=True)


def get_client_secret():
    return get_env("SHOPIFY_CLIENT_SECRET", required=True)


def get_app_url():
    return get_env("SHOPIFY_APP_URL", required=True).rstrip("/")


def get_redirect_uri():
    return get_env("SHOPIFY_REDIRECT_URI", f"{get_app_url()}/auth/callback")


def get_shopify_scopes():
    return get_env("SHOPIFY_APP_SCOPES", "")


def get_agency_api_key():
    return get_env("AGENCY_API_KEY", required=True)


def get_api_version():
    return get_env("SHOPIFY_API_VERSION", DEFAULT_SHOPIFY_API_VERSION)


def is_secure_cookie():
    return get_app_url().startswith("https://")


def get_db_path():
    return get_env(
        "OAUTH_DB_PATH",
        os.path.join(os.path.dirname(__file__), "shopify_oauth.db"),
    )


def get_db_connection():
    connection = sqlite3.connect(get_db_path())
    connection.row_factory = sqlite3.Row
    return connection


def init_db():
    with get_db_connection() as connection:
        connection.execute(
            """
            CREATE TABLE IF NOT EXISTS shop_tokens (
                shop TEXT PRIMARY KEY,
                access_token TEXT,
                scope TEXT,
                installed_at TEXT,
                updated_at TEXT NOT NULL,
                uninstalled_at TEXT
            )
            """
        )


@app.on_event("startup")
def startup_event():
    init_db()


def sanitize_shop(shop):
    shop = (shop or "").strip().lower()
    if not shop:
        raise HTTPException(status_code=400, detail="Missing shop parameter.")
    if not shop.endswith(".myshopify.com"):
        shop = f"{shop}.myshopify.com"
    allowed = set("abcdefghijklmnopqrstuvwxyz0123456789-.")
    if any(char not in allowed for char in shop):
        raise HTTPException(status_code=400, detail="Invalid Shopify shop domain.")
    if not shop.endswith(".myshopify.com"):
        raise HTTPException(status_code=400, detail="Invalid Shopify shop domain.")
    return shop


def parse_query_string(query_string):
    return urllib.parse.parse_qsl(query_string, keep_blank_values=True)


def verify_oauth_hmac(query_string, secret):
    pairs = parse_query_string(query_string)
    provided_hmac = None
    filtered_pairs = []

    for key, value in pairs:
        if key == "hmac":
            provided_hmac = value
        else:
            filtered_pairs.append((key, value))

    if not provided_hmac:
        return False

    filtered_pairs.sort(key=lambda item: item[0])
    message = "&".join(f"{key}={value}" for key, value in filtered_pairs)
    digest = hmac.new(secret.encode("utf-8"), message.encode("utf-8"), hashlib.sha256)
    computed_hmac = digest.hexdigest()
    return hmac.compare_digest(computed_hmac, provided_hmac)


def verify_webhook_hmac(raw_body, received_hmac, secret):
    digest = hmac.new(secret.encode("utf-8"), raw_body, hashlib.sha256).digest()
    encoded = base64.b64encode(digest).decode("utf-8")
    return hmac.compare_digest(encoded, received_hmac or "")


def sign_state(value):
    digest = hmac.new(
        get_client_secret().encode("utf-8"),
        value.encode("utf-8"),
        hashlib.sha256,
    )
    return f"{value}.{digest.hexdigest()}"


def unsign_state(value):
    raw_value, _, provided_signature = value.rpartition(".")
    if not raw_value or not provided_signature:
        return None
    expected_signature = hmac.new(
        get_client_secret().encode("utf-8"),
        raw_value.encode("utf-8"),
        hashlib.sha256,
    ).hexdigest()
    if not hmac.compare_digest(expected_signature, provided_signature):
        return None
    return raw_value


def get_shop_record(shop):
    with get_db_connection() as connection:
        row = connection.execute(
            """
            SELECT shop, access_token, scope, installed_at, updated_at, uninstalled_at
            FROM shop_tokens
            WHERE shop = ?
            """,
            (shop,),
        ).fetchone()
    return dict(row) if row else None


def save_shop_token(shop, access_token, scope):
    current_time = utcnow_iso()
    existing = get_shop_record(shop)
    installed_at = existing["installed_at"] if existing and existing["installed_at"] else current_time

    with get_db_connection() as connection:
        connection.execute(
            """
            INSERT INTO shop_tokens (
                shop, access_token, scope, installed_at, updated_at, uninstalled_at
            )
            VALUES (?, ?, ?, ?, ?, NULL)
            ON CONFLICT(shop) DO UPDATE SET
                access_token = excluded.access_token,
                scope = excluded.scope,
                installed_at = excluded.installed_at,
                updated_at = excluded.updated_at,
                uninstalled_at = NULL
            """,
            (shop, access_token, scope, installed_at, current_time),
        )


def mark_shop_uninstalled(shop):
    with get_db_connection() as connection:
        connection.execute(
            """
            UPDATE shop_tokens
            SET access_token = NULL, uninstalled_at = ?, updated_at = ?
            WHERE shop = ?
            """,
            (utcnow_iso(), utcnow_iso(), shop),
        )


def require_agency_api_access(request):
    authorization = request.headers.get("Authorization", "")
    prefix = "Bearer "
    if not authorization.startswith(prefix):
        raise HTTPException(status_code=401, detail="Missing bearer token.")

    provided_token = authorization[len(prefix):].strip()
    if not hmac.compare_digest(provided_token, get_agency_api_key()):
        raise HTTPException(status_code=403, detail="Invalid agency API key.")


def build_authorize_url(shop, state):
    query = {
        "client_id": get_client_id(),
        "redirect_uri": get_redirect_uri(),
        "state": state,
    }
    scopes = get_shopify_scopes()
    if scopes:
        query["scope"] = scopes

    return (
        f"https://{shop}/admin/oauth/authorize?"
        f"{urllib.parse.urlencode(query, safe=',:/')}"
    )


def exchange_code_for_token(shop, code):
    response = requests.post(
        f"https://{shop}/admin/oauth/access_token",
        headers={
            "Content-Type": "application/x-www-form-urlencoded",
            "Accept": "application/json",
        },
        data={
            "client_id": get_client_id(),
            "client_secret": get_client_secret(),
            "code": code,
        },
        timeout=30,
    )
    response.raise_for_status()
    return response.json()


@app.get("/health")
def health():
    return {"ok": True, "api_version": get_api_version()}


@app.get("/", response_class=HTMLResponse)
def root(request: Request):
    shop = request.query_params.get("shop")
    raw_query = request.url.query

    if "hmac=" in raw_query and not verify_oauth_hmac(raw_query, get_client_secret()):
        raise HTTPException(status_code=400, detail="Invalid Shopify HMAC signature.")

    if not shop:
        return HTMLResponse(
            """
            <html>
              <body>
                <h1>Shopify Bulk Tool OAuth Backend</h1>
                <p>Use this app URL as the Shopify app URL in the Dev Dashboard.</p>
                <p>To connect a store manually, visit <code>/auth/start?shop=store-name.myshopify.com</code>.</p>
              </body>
            </html>
            """
        )

    normalized_shop = sanitize_shop(shop)
    record = get_shop_record(normalized_shop)
    if record and record.get("access_token") and not record.get("uninstalled_at"):
        return HTMLResponse(
            f"""
            <html>
              <body>
                <h1>Store Connected</h1>
                <p><strong>{normalized_shop}</strong> is already connected.</p>
                <p>You can now use the desktop bulk tool against this store.</p>
              </body>
            </html>
            """
        )

    return RedirectResponse(
        url=f"/auth/start?shop={urllib.parse.quote(normalized_shop, safe='')}",
        status_code=302,
    )


@app.get("/auth/start")
def auth_start(shop: str):
    normalized_shop = sanitize_shop(shop)
    state = secrets.token_urlsafe(24)
    signed_state = sign_state(state)
    authorize_url = build_authorize_url(normalized_shop, state)

    response = RedirectResponse(url=authorize_url, status_code=302)
    response.set_cookie(
        key=STATE_COOKIE_NAME,
        value=signed_state,
        httponly=True,
        secure=is_secure_cookie(),
        samesite="lax",
        max_age=600,
    )
    return response


@app.get("/auth/callback", response_class=HTMLResponse)
def auth_callback(request: Request):
    raw_query = request.url.query
    if not verify_oauth_hmac(raw_query, get_client_secret()):
        raise HTTPException(status_code=400, detail="Invalid Shopify callback HMAC.")

    shop = sanitize_shop(request.query_params.get("shop"))
    code = request.query_params.get("code")
    state = request.query_params.get("state")
    signed_cookie_state = request.cookies.get(STATE_COOKIE_NAME)

    if not code or not state or not signed_cookie_state:
        raise HTTPException(status_code=400, detail="Missing OAuth callback parameters.")

    original_state = unsign_state(signed_cookie_state)
    if not original_state or not hmac.compare_digest(original_state, state):
        raise HTTPException(status_code=400, detail="Invalid OAuth state.")

    token_payload = exchange_code_for_token(shop, code)
    access_token = token_payload.get("access_token")
    scope = token_payload.get("scope", "")
    if not access_token:
        raise HTTPException(status_code=502, detail="Shopify did not return an access token.")

    save_shop_token(shop, access_token, scope)

    response = HTMLResponse(
        f"""
        <html>
          <body>
            <h1>Store Connected</h1>
            <p><strong>{shop}</strong> is now connected.</p>
            <p>The desktop bulk tool can now fetch its token from the OAuth backend.</p>
          </body>
        </html>
        """
    )
    response.delete_cookie(STATE_COOKIE_NAME)
    return response


@app.get("/api/shops")
def list_connected_shops(request: Request):
    require_agency_api_access(request)

    with get_db_connection() as connection:
        rows = connection.execute(
            """
            SELECT shop, scope, installed_at, updated_at, uninstalled_at
            FROM shop_tokens
            ORDER BY shop ASC
            """
        ).fetchall()

    return {
        "shops": [dict(row) for row in rows],
        "api_version": get_api_version(),
    }


@app.get("/api/shops/{shop}/token")
def get_shop_token(shop: str, request: Request):
    require_agency_api_access(request)
    normalized_shop = sanitize_shop(shop)
    record = get_shop_record(normalized_shop)
    if not record or not record.get("access_token") or record.get("uninstalled_at"):
        raise HTTPException(
            status_code=404,
            detail=f"No active Shopify token stored for {normalized_shop}.",
        )

    return JSONResponse(
        {
            "shop": normalized_shop,
            "access_token": record["access_token"],
            "scope": record.get("scope", ""),
            "api_version": get_api_version(),
            "updated_at": record.get("updated_at"),
        }
    )


@app.post("/webhooks/app/uninstalled")
async def app_uninstalled_webhook(request: Request):
    raw_body = await request.body()
    received_hmac = request.headers.get("X-Shopify-Hmac-Sha256")

    if not verify_webhook_hmac(raw_body, received_hmac, get_client_secret()):
        raise HTTPException(status_code=401, detail="Invalid webhook signature.")

    shop_domain = request.headers.get("X-Shopify-Shop-Domain")
    if not shop_domain:
        raise HTTPException(status_code=400, detail="Missing Shopify shop domain header.")

    normalized_shop = sanitize_shop(shop_domain)
    mark_shop_uninstalled(normalized_shop)
    return {"ok": True, "shop": normalized_shop}
