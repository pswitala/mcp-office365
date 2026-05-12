import { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import { mkdir, readFile, writeFile } from "fs/promises";
import { existsSync } from "fs";
import { join } from "path";
import { homedir } from "os";

const cacheDir = join(homedir(), ".mcp-o365");
const cachePath = join(cacheDir, "token.json");

async function ensureDir(): Promise<void> {
  if (!existsSync(cacheDir)) {
    await mkdir(cacheDir, { recursive: true });
  }
}

export const tokenCachePlugin: ICachePlugin = {
  async beforeCacheAccess(ctx: TokenCacheContext): Promise<void> {
    await ensureDir();
    if (existsSync(cachePath)) {
      const data = await readFile(cachePath, "utf-8");
      ctx.tokenCache.deserialize(data);
    }
  },

  async afterCacheAccess(ctx: TokenCacheContext): Promise<void> {
    if (ctx.cacheHasChanged) {
      await ensureDir();
      await writeFile(cachePath, ctx.tokenCache.serialize(), "utf-8");
    }
  },
};
