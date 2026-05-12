import { PublicClientApplication, DeviceCodeRequest } from "@azure/msal-node";
import { saveAccount, type AccountEntry } from "./accounts.js";
import { getMsalApp } from "./msal.js";
import { GRAPH_SCOPES } from "../constants.js";

export interface LoginResult {
  success: boolean;
  alias?: string;
  username?: string;
  message: string;
}

export async function performLogin(
  app: PublicClientApplication,
  alias: string
): Promise<LoginResult> {
  const req: DeviceCodeRequest = {
    scopes: GRAPH_SCOPES,
    deviceCodeCallback: (response) => {
      console.log("\n" + response.message + "\n");
    },
  };

  try {
    const result = await app.acquireTokenByDeviceCode(req);
    if (!result?.account) throw new Error("No account returned.");

    await saveAccount(alias, {
      homeAccountId: result.account.homeAccountId,
      username: result.account.username,
      tenantId: result.account.tenantId ?? "",
    });

    return {
      success: true,
      alias,
      username: result.account.username,
      message: `Logged in as ${result.account.username} (alias: "${alias}"). Token saved to ~/.mcp-o365/.`,
    };
  } catch (err) {
    return {
      success: false,
      alias,
      message: `Login failed: ${err instanceof Error ? err.message : String(err)}`,
    };
  }
}
