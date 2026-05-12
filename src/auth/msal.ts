import {
  PublicClientApplication,
  Configuration,
  AccountInfo,
  SilentFlowRequest,
} from "@azure/msal-node";
import { tokenCachePlugin } from "./tokenCache.js";
import { getAccount } from "./accounts.js";
import { GRAPH_SCOPES } from "../constants.js";
import { getClientId } from "../config.js";

let _app: PublicClientApplication | null = null;

export async function getMsalApp(): Promise<PublicClientApplication> {
  if (_app) return _app;

  const clientId = await getClientId();
  if (!clientId) {
    throw new Error(
      "No Office 365 client ID configured.\n" +
      "Use the 'configure' tool to set your Azure AD application client ID,\n" +
      "or set the MCP_O365_CLIENT_ID environment variable."
    );
  }

  const config: Configuration = {
    auth: {
      clientId,
      authority: "https://login.microsoftonline.com/common",
    },
    cache: {
      cachePlugin: tokenCachePlugin,
    },
  };

  _app = new PublicClientApplication(config);
  return _app;
}

export async function acquireTokenForAlias(alias: string): Promise<string> {
  const app = await getMsalApp();
  const entry = await getAccount(alias);

  if (entry) {
    const accounts = await app.getTokenCache().getAllAccounts();
    const account = accounts.find((a) => a.homeAccountId === entry.homeAccountId);

    if (account) {
      const req: SilentFlowRequest = { scopes: GRAPH_SCOPES, account };
      try {
        const result = await app.acquireTokenSilent(req);
        if (result?.accessToken) return result.accessToken;
      } catch {
        // Fall through to device code
      }
    }
  }

  throw new Error(
    `No valid token for alias "${alias}". Run the login tool first: { "alias": "${alias}" }`
  );
}

