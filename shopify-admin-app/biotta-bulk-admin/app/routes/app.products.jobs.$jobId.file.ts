import type { LoaderFunctionArgs } from "react-router";

import { buildProductJobFileResponse } from "../product-bulk.server";
import { authenticate } from "../shopify.server";

export const loader = async ({ params, request }: LoaderFunctionArgs) => {
  try {
    await authenticate.admin(request);

    const jobId = params.jobId;

    if (!jobId) {
      throw new Error("A product job id is required.");
    }

    return await buildProductJobFileResponse(jobId);
  } catch (error) {
    if (error instanceof Response) {
      throw error;
    }

    return Response.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Could not download the completed spreadsheet.",
      },
      { status: 500 },
    );
  }
};
