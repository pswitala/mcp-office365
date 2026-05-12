#!/usr/bin/env node
/**
 * CLI wrapper for auth operations.
 *   node dist/login.js --alias work       # login
 *   node dist/login.js list               # list aliases
 *   node dist/login.js remove <alias>     # remove alias
 */
import { getMsalApp } from "./auth/msal.js";
import { listAliases, removeAccount } from "./auth/accounts.js";
import { performLogin } from "./auth/login.js";
import { getClientId } from "./config.js";

const args = process.argv.slice(2);
const cmd = args[0];

if (cmd === "list") {
  const aliases = await listAliases();
  if (aliases.length === 0) {
    console.log("No mailboxes configured.");
  } else {
    console.log("Configured aliases:");
    for (const a of aliases) console.log(`  - ${a}`);
  }
  process.exit(0);
}

if (cmd === "remove") {
  const alias = args[1];
  if (!alias) { console.error("Usage: login.js remove <alias>"); process.exit(1); }
  await removeAccount(alias);
  console.log(`Removed alias "${alias}".`);
  process.exit(0);
}

// Default: add/login
const aliasFlag = args.indexOf("--alias");
const alias = aliasFlag >= 0 ? args[aliasFlag + 1] : args[0];

if (!alias) {
  console.error([
    "Usage:",
    "  node dist/login.js --alias <name>    # add a mailbox",
    "  node dist/login.js list              # list aliases",
    "  node dist/login.js remove <alias>   # remove an alias",
  ].join("\n"));
  process.exit(1);
}

const clientId = await getClientId();
if (!clientId) {
  console.error("ERROR: No client ID configured.");
  console.error("Set MCP_O365_CLIENT_ID env var or use the 'configure' MCP tool.");
  process.exit(1);
}

console.log(`Starting login for alias "${alias}"...`);

const app = await getMsalApp();
const result = await performLogin(app, alias);
console.log(result.message);
process.exit(result.success ? 0 : 1);
