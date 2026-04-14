import "@shopify/shopify-app-react-router/adapters/node";
import {
  ApiVersion,
  AppDistribution,
  shopifyApp,
} from "@shopify/shopify-app-react-router/server";
import { PrismaSessionStorage } from "@shopify/shopify-app-session-storage-prisma";
import prisma from "./db.server";

const prismaSessionStorage = new PrismaSessionStorage(prisma) as unknown as NonNullable<
  Parameters<typeof shopifyApp>[0]["sessionStorage"]
>;

const shopify = shopifyApp({
  apiKey: process.env.SHOPIFY_API_KEY,
  apiSecretKey: process.env.SHOPIFY_API_SECRET || "",
  apiVersion: ApiVersion.October25,
  scopes: process.env.SCOPES?.split(","),
  appUrl: process.env.SHOPIFY_APP_URL || process.env.RENDER_EXTERNAL_URL || "",
  authPathPrefix: "/auth",
  // The scaffold currently installs mismatched Shopify package versions, so we
  // normalize the adapter type here until the dependencies are pinned together.
  sessionStorage: prismaSessionStorage,
  distribution: AppDistribution.SingleMerchant,
  future: {
    expiringOfflineAccessTokens: true,
  },
  ...(process.env.SHOP_CUSTOM_DOMAIN
    ? { customShopDomains: [process.env.SHOP_CUSTOM_DOMAIN] }
    : {}),
});

export default shopify;
export const apiVersion = ApiVersion.October25;
export const addDocumentResponseHeaders = shopify.addDocumentResponseHeaders;
export const authenticate = shopify.authenticate;
export const unauthenticated = shopify.unauthenticated;
export const login = shopify.login;
export const registerWebhooks = shopify.registerWebhooks;
export const sessionStorage = shopify.sessionStorage;

type AuthTraceDetails = {
  hasAuthorizationHeader: boolean;
  hasHost: boolean;
  hasIdToken: boolean;
  isEmbedded: boolean;
  path: string;
  routeId?: string;
  searchKeys: string[];
  shop: string | null;
  userAgent: string;
};

function getAuthTraceDetails(request: Request, routeId?: string): AuthTraceDetails {
  const url = new URL(request.url);
  const userAgent = request.headers.get("user-agent") ?? "";

  return {
    routeId,
    path: url.pathname,
    searchKeys: Array.from(url.searchParams.keys()).sort(),
    hasIdToken: url.searchParams.has("id_token"),
    hasAuthorizationHeader: Boolean(request.headers.get("authorization")),
    hasHost: Boolean(url.searchParams.get("host")),
    isEmbedded: url.searchParams.get("embedded") === "1",
    shop: url.searchParams.get("shop"),
    userAgent: userAgent.slice(0, 180),
  };
}

export async function authenticateAdminWithTrace(
  request: Request,
  routeId?: string,
) {
  const details = getAuthTraceDetails(request, routeId);
  console.log("[auth-trace] start", JSON.stringify(details));

  try {
    const context = await authenticate.admin(request);
    console.log(
      "[auth-trace] success",
      JSON.stringify({
        routeId,
        path: details.path,
        hasIdToken: details.hasIdToken,
        hasAuthorizationHeader: details.hasAuthorizationHeader,
        shop: details.shop,
      }),
    );
    return context;
  } catch (error) {
    if (error instanceof Response) {
      console.log(
        "[auth-trace] response",
        JSON.stringify({
          routeId,
          path: details.path,
          hasIdToken: details.hasIdToken,
          hasAuthorizationHeader: details.hasAuthorizationHeader,
          status: error.status,
          statusText: error.statusText,
          location: error.headers.get("location"),
          shop: details.shop,
        }),
      );
    } else {
      console.error("[auth-trace] error", {
        routeId,
        path: details.path,
        shop: details.shop,
        error,
      });
    }

    throw error;
  }
}
