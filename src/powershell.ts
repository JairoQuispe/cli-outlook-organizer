import { getScriptPath } from "./embedded-scripts";
import type { PstFolder, Store } from "./types";

export type ProcessEvent =
  | { type: "stdout"; line: string }
  | { type: "stderr"; line: string }
  | { type: "exit"; code: number };

export async function runPreflight(): Promise<string[]> {
  const messages: string[] = [];
  if (process.platform !== "win32") {
    throw new Error("Esta CLI solo esta soportada en Windows.");
  }

  const checks = [
    {
      label: "Outlook clasico instalado",
      command: "try { $null = New-Object -ComObject Outlook.Application; exit 0 } catch { exit 1 }",
    },
    {
      label: "Outlook abierto y con proceso activo",
      command: "if (Get-Process OUTLOOK -ErrorAction SilentlyContinue) { exit 0 } else { exit 1 }",
    },
  ];

  for (const check of checks) {
    const result = await Bun.spawn(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-Command", check.command]).exited;
    if (result !== 0) throw new Error(check.label);
    messages.push(check.label);
  }

  return messages;
}

export async function listStores(): Promise<Store[]> {
  const scriptPath = await getScriptPath("list-stores");
  const proc = Bun.spawn(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-Json"], {
    stdout: "pipe",
    stderr: "pipe",
  });

  const [stdout, stderr, code] = await Promise.all([new Response(proc.stdout).text(), new Response(proc.stderr).text(), proc.exited]);
  if (code !== 0) throw new Error(stderr.trim() || "No se pudo listar los buzones de Outlook.");

  const jsonLine = stdout
    .split(/\r?\n/)
    .map((line) => line.trim())
    .find((line) => line.startsWith("{"));

  if (!jsonLine) throw new Error("El script de buzones no devolvio JSON valido.");
  const payload = JSON.parse(jsonLine) as { stores?: Store[] };
  return payload.stores ?? [];
}

export async function listPstFolders(pstPath: string, onEvent: (event: ProcessEvent) => void): Promise<PstFolder[]> {
  const scriptPath = await getScriptPath("import-pst");
  const events = await collectProcess(["powershell", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", scriptPath, "-PstPath", pstPath, "-ListFolders", "-Json"], onEvent);
  const folders: PstFolder[] = [];

  for (const line of events.stdout) {
    if (!line.trim().startsWith("{")) continue;
    const payload = JSON.parse(line) as Record<string, unknown>;
    if (payload.type === "folder") {
      folders.push({
        path: String(payload.path ?? ""),
        itemCount: Number(payload.itemCount ?? 0),
        yearSummary: payload.yearSummary ? String(payload.yearSummary) : undefined,
      });
    }
  }

  return folders;
}

export async function executeImport(options: {
  pstPath: string;
  targetStoreId: string;
  action: "Copy" | "Move";
  includeFolders: string[];
  filterYear?: string;
  filterMonths?: string;
  skipDuplicates: boolean;
  onEvent: (event: ProcessEvent) => void;
}): Promise<number> {
  const scriptPath = await getScriptPath("import-pst");
  const args = [
    "powershell",
    "-NoProfile",
    "-ExecutionPolicy",
    "Bypass",
    "-File",
    scriptPath,
    "-PstPath",
    options.pstPath,
    "-TargetStoreId",
    options.targetStoreId,
    "-Action",
    options.action,
    "-IncludeFoldersJson",
    JSON.stringify(options.includeFolders),
    "-Json",
  ];

  if (options.filterYear) args.push("-FilterOnlyYear", options.filterYear);
  if (options.filterMonths) args.push("-FilterOnlyMonths", options.filterMonths);
  if (options.skipDuplicates) args.push("-SkipDuplicates");

  const result = await collectProcess(args, options.onEvent);
  return result.code;
}

async function collectProcess(args: string[], onEvent: (event: ProcessEvent) => void): Promise<{ stdout: string[]; stderr: string[]; code: number }> {
  const proc = Bun.spawn(args, { stdout: "pipe", stderr: "pipe" });
  const stdoutPromise = readLines(proc.stdout, (line) => onEvent({ type: "stdout", line }));
  const stderrPromise = readLines(proc.stderr, (line) => onEvent({ type: "stderr", line }));
  const [stdout, stderr, code] = await Promise.all([stdoutPromise, stderrPromise, proc.exited]);
  onEvent({ type: "exit", code });
  return { stdout, stderr, code };
}

async function readLines(stream: ReadableStream<Uint8Array>, onLine: (line: string) => void): Promise<string[]> {
  const text = await new Response(stream).text();
  const lines = text.split(/\r?\n/).filter((line) => line.length > 0);
  for (const line of lines) onLine(line);
  return lines;
}
