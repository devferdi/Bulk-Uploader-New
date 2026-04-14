import type { ActionFunctionArgs, LoaderFunctionArgs } from "react-router";

import { startMetaobjectUploadJob } from "../product-bulk.server";
import { authenticate } from "../shopify.server";

export const loader = async ({ request }: LoaderFunctionArgs) => {
  return Response.json({ error: "Use the upload form." }, { status: 405 });
};

export const action = async ({ request }: ActionFunctionArgs) => {
  try {
    const { session } = await authenticate.admin(request);
    const accessToken = session.accessToken;
    const formData = await request.formData();
    const spreadsheet = formData.get("spreadsheet");

    if (!accessToken) {
      throw new Error("No Shopify access token is available for this shop.");
    }

    if (!(spreadsheet instanceof File)) {
      throw new Error("Choose an .xlsx spreadsheet before uploading.");
    }

    return Response.json(
      await startMetaobjectUploadJob({
        accessToken,
        file: spreadsheet,
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
            : "Could not start the metaobjects upload job.",
      },
      { status: 500 },
    );
  }
};
