import { mkdir, readFile, writeFile } from "fs/promises";
import { existsSync } from "fs";
import { join } from "path";
import { homedir } from "os";

const cacheDir = join(homedir(), ".mcp-o365");
const configPath = join(cacheDir, "config.json");

export interface Config {
  clientId: string;
}

async function ensureDir(): Promise<void> {
  if (!existsSync(cacheDir)) {
    await mkdir(cacheDir, { recursive: true });
  }
}

export async function saveConfig(clientId: string): Promise<void> {
  await ensureDir();
  await writeFile(configPath, JSON.stringify({ clientId }, null, 2), "utf-8");
}

export async function loadConfig(): Promise<Config | null> {
  if (!existsSync(configPath)) return null;
  const raw = await readFile(configPath, "utf-8");
  return JSON.parse(raw) as Config;
}

export async function getClientId(): Promise<string | null> {
  const envId = process.env["MCP_O365_CLIENT_ID"];
  if (envId) return envId;
  const cfg = await loadConfig();
  return cfg?.clientId ?? null;
}
