import type { LoaderFunctionArgs } from "react-router";

export const loader = async (_args: LoaderFunctionArgs) => {
  return Response.json({ ok: true });
};

