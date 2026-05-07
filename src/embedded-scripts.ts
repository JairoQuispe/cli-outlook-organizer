import listStoresScript from "../scripts-powershell/outlook-list-stores.ps1" with { type: "text" };
import importPstScript from "../scripts-powershell/outlook-import-pst.ps1" with { type: "text" };
import { mkdir, rm } from "node:fs/promises";

export type ScriptKind = "list-stores" | "import-pst";

const scripts: Record<ScriptKind, { fileName: string; content: string }> = {
  "list-stores": {
    fileName: "outlook-list-stores.ps1",
    content: listStoresScript,
  },
  "import-pst": {
    fileName: "outlook-import-pst.ps1",
    content: importPstScript,
  },
};

const extractedPaths = new Map<ScriptKind, string>();
let extractionDir: string | undefined;

export async function getScriptPath(kind: ScriptKind): Promise<string> {
  const existing = extractedPaths.get(kind);
  if (existing) return existing;

  extractionDir ??= await createExtractionDir();
  const script = scripts[kind];
  const path = `${extractionDir}\\${script.fileName}`;
  await Bun.write(path, script.content);
  extractedPaths.set(kind, path);
  return path;
}

export async function cleanupScripts(): Promise<void> {
  if (!extractionDir) return;
  await rm(extractionDir, { recursive: true, force: true });
  extractedPaths.clear();
  extractionDir = undefined;
}

async function createExtractionDir(): Promise<string> {
  const temp = process.env.TEMP || "C:\\Windows\\Temp";
  const dir = `${temp}\\outlook-organizer-open-tui-${process.pid}`;
  await mkdir(dir, { recursive: true });
  return dir;
}
