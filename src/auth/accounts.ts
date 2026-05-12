import { mkdir, readFile, writeFile } from "fs/promises";
import { existsSync } from "fs";
import { join } from "path";
import { homedir } from "os";

const cacheDir = join(homedir(), ".mcp-o365");
const accountsPath = join(cacheDir, "accounts.json");

export interface AccountEntry {
  homeAccountId: string;
  username: string;
  tenantId: string;
}

type AccountMap = Record<string, AccountEntry>;

async function ensureDir(): Promise<void> {
  if (!existsSync(cacheDir)) {
    await mkdir(cacheDir, { recursive: true });
  }
}

async function load(): Promise<AccountMap> {
  if (!existsSync(accountsPath)) return {};
  const raw = await readFile(accountsPath, "utf-8");
  return JSON.parse(raw) as AccountMap;
}

async function save(map: AccountMap): Promise<void> {
  await ensureDir();
  await writeFile(accountsPath, JSON.stringify(map, null, 2), "utf-8");
}

export async function listAliases(): Promise<string[]> {
  const map = await load();
  return Object.keys(map);
}

export async function getAccount(alias: string): Promise<AccountEntry | undefined> {
  const map = await load();
  return map[alias];
}

export async function saveAccount(alias: string, entry: AccountEntry): Promise<void> {
  const map = await load();
  map[alias] = entry;
  await save(map);
}

export async function removeAccount(alias: string): Promise<void> {
  const map = await load();
  delete map[alias];
  await save(map);
}
