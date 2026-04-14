
import type { HeadersFunction, LoaderFunctionArgs } from "react-router";
import { authenticateAdminWithTrace } from "../shopify.server";
import { boundary } from "@shopify/shopify-app-react-router/server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  await authenticateAdminWithTrace(request, "routes/auth.$");

  return null;
};

export const headers: HeadersFunction = (headersArgs) => {
  return boundary.headers(headersArgs);
};
