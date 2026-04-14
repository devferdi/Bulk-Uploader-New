import type { LoaderFunctionArgs } from "react-router";

import { getProductJobSummary } from "../product-bulk.server";
import { authenticate } from "../shopify.server";

export const loader = async ({ params, request }: LoaderFunctionArgs) => {
  try {
    await authenticate.admin(request);

    const jobId = params.jobId;

    if (!jobId) {
      throw new Error("A product job id is required.");
    }

    const jobSummary = await getProductJobSummary(jobId);

    if (!jobSummary) {
      return Response.json(
        { error: "This product spreadsheet job was not found or has expired." },
        { status: 404 },
      );
    }

    return Response.json(jobSummary);
  } catch (error) {
    if (error instanceof Response) {
      throw error;
    }

    return Response.json(
      {
        error:
          error instanceof Error
            ? error.message
            : "Could not load the product job status.",
      },
      { status: 500 },
    );
  }
};
