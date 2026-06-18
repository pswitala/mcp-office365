import { Client, GraphRequest } from "@microsoft/microsoft-graph-client";
import { acquireTokenForAlias } from "../auth/msal.js";

async function withRetry<T>(fn: () => Promise<T>): Promise<T> {
  try {
    return await fn();
  } catch (err: unknown) {
    const e = err as { statusCode?: number; headers?: Record<string, string> };
    if (e?.statusCode === 429) {
      const retryAfter = parseInt(e.headers?.["retry-after"] ?? "5", 10);
      await new Promise((r) => setTimeout(r, retryAfter * 1000));
      return await fn();
    }
    throw err;
  }
}

export async function getGraphClient(alias: string): Promise<Client> {
  const token = await acquireTokenForAlias(alias);
  return Client.init({
    authProvider: (done) => done(null, token),
  });
}

export async function graphGet<T>(
  alias: string,
  path: string,
  opts?: { query?: Record<string, string | number>; headers?: Record<string, string> }
): Promise<T> {
  const client = await getGraphClient(alias);
  return withRetry(() => {
    let req = client.api(path);
    if (opts?.query) req = req.query(opts.query);
    if (opts?.headers) {
      for (const [k, v] of Object.entries(opts.headers)) req = req.header(k, v);
    }
    return req.get() as Promise<T>;
  });
}

export async function graphPost<T>(alias: string, path: string, body: unknown): Promise<T> {
  const client = await getGraphClient(alias);
  return withRetry(() => client.api(path).post(body) as Promise<T>);
}

export async function graphPatch<T>(alias: string, path: string, body: unknown): Promise<T> {
  const client = await getGraphClient(alias);
  return withRetry(() => client.api(path).patch(body) as Promise<T>);
}

export async function graphDelete(alias: string, path: string): Promise<void> {
  const client = await getGraphClient(alias);
  return withRetry(() => client.api(path).delete() as Promise<void>);
}
