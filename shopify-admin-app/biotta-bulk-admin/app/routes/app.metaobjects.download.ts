import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";

import { startMetaobjectDownloadJob } from "../product-bulk.server";
import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  return Response.json({ error: "Use POST to start a metaobjects download job." }, { status: 405 });
};

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    const { session } = await authenticate.admin(request);
    const accessToken = session.accessToken;

    if (!accessToken) {
      throw new Error("No Shopify access token is available for this shop.");
    }

    return Response.json(
      await startMetaobjectDownloadJob({
        accessToken,
        shop: session.shop,
      }),
    );
  } catch (error) {
    if (error instanceof Response) {
      throw error;
    }

    return Response.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Could not start the metaobjects download job.",
      },
      { status: 500 },
    );
  }
};
