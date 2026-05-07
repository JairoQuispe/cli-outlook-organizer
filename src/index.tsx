/** @jsxImportSource @opentui/react */
import { createCliRenderer } from "@opentui/core";
import { createDefaultOpenTuiKeymap } from "@opentui/keymap/opentui";
import { KeymapProvider, useBindings } from "@opentui/keymap/react";
import { createRoot, useKeyboard, useRenderer } from "@opentui/react";
import { useEffect, useMemo, useState } from "react";
import { cleanupScripts } from "./embedded-scripts";
import { executeImport, listPstFolders, listStores, runPreflight, type ProcessEvent } from "./powershell";
import type { ActionId, AppScreen, ImportOptions, PstFolder, Store } from "./types";

const actions: Array<{ id: ActionId; label: string; description: string; enabled: boolean }> = [
  { id: "import-pst", label: "Importar correos desde PST", description: "Wizard OpenTUI para importar carpetas seleccionadas", enabled: true },
  { id: "backup-pst", label: "Respaldar correos hacia PST", description: "Omitido: script no usado por el flujo activo", enabled: false },
];

function App() {
  const renderer = useRenderer();
  const [screen, setScreen] = useState<AppScreen>("preflight");
  const [steps, setSteps] = useState<string[]>([]);
  const [error, setError] = useState<string>();
  const [stores, setStores] = useState<Store[]>([]);
  const [storeIndex, setStoreIndex] = useState(0);
  const [actionIndex, setActionIndex] = useState(0);
  const [selectedStore, setSelectedStore] = useState<Store>();
  const [pstPath, setPstPath] = useState("");
  const [options, setOptions] = useState<ImportOptions>({ action: "Copy", skipDuplicates: true });
  const [optionIndex, setOptionIndex] = useState(0);
  const [folders, setFolders] = useState<PstFolder[]>([]);
  const [folderIndex, setFolderIndex] = useState(0);
  const [selectedFolders, setSelectedFolders] = useState<Set<string>>(new Set());
  const [log, setLog] = useState<string[]>(["Iniciando Outlook Organizer con OpenTUI..."]);

  useBindings(
    () => ({
      commands: [{ name: "quit", run: () => renderer.destroy() }],
      bindings: [{ key: "q", cmd: "quit" }, { key: "ctrl+c", cmd: "quit" }],
    }),
    [renderer],
  );

  useEffect(() => {
    let cancelled = false;
    async function boot() {
      try {
        setLog((prev) => [...prev, "Ejecutando preflight de Windows y Outlook..."]);
        const preflightSteps = await runPreflight();
        if (cancelled) return;
        setSteps(preflightSteps);
        setLog((prev) => [...prev, "Listando buzones conectados a Outlook..."]);
        const found = (await listStores()).filter((store) => store.storeId && store.storeId.length > 0);
        if (cancelled) return;
        if (found.length === 0) throw new Error("No hay buzones con StoreId valido para continuar.");
        setStores(found);
        setScreen("stores");
        setLog((prev) => [...prev, `${found.length} buzon(es) disponible(s).`]);
      } catch (err) {
        if (cancelled) return;
        setError(err instanceof Error ? err.message : String(err));
      }
    }
    boot();
    return () => {
      cancelled = true;
    };
  }, []);

  useKeyboard((key) => {
    if (key.name === "escape") {
      if (screen === "pst-path" || screen === "import-options" || screen === "folders") setScreen("actions");
      else renderer.destroy();
    }
    if (screen === "stores" && stores.length > 0) {
      if (key.name === "up") setStoreIndex((value) => (value === 0 ? stores.length - 1 : value - 1));
      if (key.name === "down") setStoreIndex((value) => (value + 1) % stores.length);
      if (key.name === "return") {
        setSelectedStore(stores[storeIndex]);
        setScreen("actions");
      }
    }
    if (screen === "actions") {
      if (key.name === "up") setActionIndex((value) => (value === 0 ? actions.length - 1 : value - 1));
      if (key.name === "down") setActionIndex((value) => (value + 1) % actions.length);
      if (key.name === "return" && actions[actionIndex]?.enabled) setScreen("pst-path");
    }
    if (screen === "import-options") {
      if (key.name === "up") setOptionIndex((value) => (value === 0 ? 4 : value - 1));
      if (key.name === "down") setOptionIndex((value) => (value + 1) % 5);
      if (key.name === "space" || key.name === "return") toggleOption();
    }
    if (screen === "folders" && folders.length > 0) {
      if (key.name === "up") setFolderIndex((value) => (value === 0 ? folders.length - 1 : value - 1));
      if (key.name === "down") setFolderIndex((value) => (value + 1) % folders.length);
      if (key.name === "space") toggleFolder(folders[folderIndex]);
      if (key.name === "a") selectAllImportableFolders();
      if (key.name === "n") setSelectedFolders(new Set());
      if (key.name === "return") void executeSelectedImport();
    }
  });

  async function scanFolders(path: string) {
    const trimmed = path.trim().replace(/^["']|["']$/g, "");
    if (!trimmed) return;
    setPstPath(trimmed);
    setScreen("scan");
    setError(undefined);
    setLog((prev) => [...prev, `Escaneando PST: ${trimmed}`]);
    try {
      const found = await listPstFolders(trimmed, appendProcessEvent);
      const importable = found.filter((folder) => folder.itemCount > 0);
      if (importable.length === 0) throw new Error("El PST no contiene carpetas importables.");
      setFolders(found);
      setSelectedFolders(new Set(importable.map((folder) => folder.path)));
      setFolderIndex(0);
      setScreen("folders");
      setLog((prev) => [...prev, `${importable.length} carpeta(s) con correos seleccionadas.`]);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
      setScreen("pst-path");
    }
  }

  function toggleOption() {
    if (optionIndex === 0) setOptions((prev) => ({ ...prev, action: prev.action === "Copy" ? "Move" : "Copy" }));
    if (optionIndex === 1) setOptions((prev) => ({ ...prev, skipDuplicates: !prev.skipDuplicates }));
    if (optionIndex === 2) setOptions((prev) => ({ ...prev, filterYear: prev.filterYear ? undefined : new Date().getFullYear().toString() }));
    if (optionIndex === 3) void scanFolders(pstPath);
    if (optionIndex === 4) setScreen("pst-path");
  }

  function toggleFolder(folder: PstFolder) {
    if (folder.itemCount <= 0) return;
    setSelectedFolders((prev) => {
      const next = new Set(prev);
      if (next.has(folder.path)) next.delete(folder.path);
      else next.add(folder.path);
      return next;
    });
  }

  function selectAllImportableFolders() {
    setSelectedFolders(new Set(folders.filter((folder) => folder.itemCount > 0).map((folder) => folder.path)));
  }

  async function executeSelectedImport() {
    if (!selectedStore?.storeId) return;
    const includeFolders = [...selectedFolders];
    if (includeFolders.length === 0) {
      setError("Selecciona al menos una carpeta con correos.");
      return;
    }
    setScreen("done");
    setError(undefined);
    setLog((prev) => [...prev, `Ejecutando importacion de ${includeFolders.length} carpeta(s)...`]);
    try {
      const code = await executeImport({
        pstPath,
        targetStoreId: selectedStore.storeId,
        action: options.action,
        includeFolders,
        filterYear: options.filterYear,
        filterMonths: options.filterMonths,
        skipDuplicates: options.skipDuplicates,
        onEvent: appendProcessEvent,
      });
      setLog((prev) => [...prev, code === 0 ? "Importacion finalizada correctamente." : `Importacion finalizo con codigo ${code}.`]);
    } catch (err) {
      setError(err instanceof Error ? err.message : String(err));
    }
  }

  function appendProcessEvent(event: ProcessEvent) {
    if (event.type !== "exit") setLog((prev) => [...prev.slice(-16), event.line]);
  }

  const selected = useMemo(() => stores[storeIndex], [stores, storeIndex]);

  return (
    <box style={{ flexDirection: "column", padding: 1, gap: 1 }}>
      <box style={{ border: true, padding: 1, flexDirection: "column" }}>
        <text fg="#67e8f9"><strong>Outlook Organizer</strong> <span fg="#94a3b8">OpenTUI + React</span></text>
        <text fg="#94a3b8">↑/↓ navegar · Espacio marcar · Enter confirmar · Esc volver/salir · q salir</text>
      </box>

      {error ? <StatusPanel title="Error" color="#f87171" lines={[error]} /> : null}
      {screen === "preflight" ? <StatusPanel title="Preflight" color="#facc15" lines={steps.length ? steps : ["Verificando Outlook..."]} /> : null}
      {screen === "stores" ? <StorePanel stores={stores} selectedIndex={storeIndex} /> : null}
      {screen === "actions" && selectedStore ? <ActionPanel selectedStore={selectedStore} selectedIndex={actionIndex} /> : null}
      {screen === "pst-path" ? <PstPathPanel pstPath={pstPath} onInput={setPstPath} onSubmit={scanFolders} /> : null}
      {screen === "import-options" ? <OptionsPanel options={options} selectedIndex={optionIndex} /> : null}
      {screen === "scan" ? <StatusPanel title="Escaneando PST" color="#facc15" lines={log.slice(-12)} /> : null}
      {screen === "folders" ? <FolderPanel folders={folders} selectedIndex={folderIndex} selectedFolders={selectedFolders} options={options} /> : null}
      {screen === "done" ? <StatusPanel title="Ejecucion" color="#22c55e" lines={log.slice(-14)} /> : null}
      {selected ? <StatusPanel title="Buzon enfocado" color="#38bdf8" lines={[selected.displayName, selected.filePath ?? "Sin archivo local"]} /> : null}
    </box>
  );
}

function StorePanel({ stores, selectedIndex }: { stores: Store[]; selectedIndex: number }) {
  return (
    <box title="Buzones disponibles" style={{ border: true, padding: 1, flexDirection: "column", gap: 0 }}>
      {stores.map((store, index) => (
        <text key={store.storeId ?? index} fg={index === selectedIndex ? "#020617" : "#e2e8f0"} bg={index === selectedIndex ? "#22d3ee" : undefined}>
          {index === selectedIndex ? " > " : "   "}{index + 1}. {store.displayName} <span fg={index === selectedIndex ? "#0f172a" : "#94a3b8"}>{store.filePath ?? "Exchange/Online"}</span>
        </text>
      ))}
    </box>
  );
}

function ActionPanel({ selectedStore, selectedIndex }: { selectedStore: Store; selectedIndex: number }) {
  return (
    <box title="Acciones" style={{ border: true, padding: 1, flexDirection: "column", gap: 1 }}>
      <text fg="#a7f3d0">Buzon: {selectedStore.displayName}</text>
      {actions.map((action, index) => (
        <text key={action.id} fg={!action.enabled ? "#64748b" : index === selectedIndex ? "#020617" : "#e2e8f0"} bg={index === selectedIndex ? "#a3e635" : undefined}>
          {index === selectedIndex ? " > " : "   "}{action.label} <span>{action.description}</span>
        </text>
      ))}
    </box>
  );
}

function PstPathPanel({ pstPath, onInput, onSubmit }: { pstPath: string; onInput: (value: string) => void; onSubmit: (value: string) => void }) {
  const submitCurrentPath = () => onSubmit(pstPath);
  return (
    <box title="Archivo PST" style={{ border: true, padding: 1, flexDirection: "column", gap: 1 }}>
      <text fg="#e2e8f0">Escribe la ruta completa del archivo .pst y presiona Enter.</text>
      <box title="Ruta" style={{ border: true, height: 3, width: 100 }}>
        <input value={pstPath} placeholder="C:\\Users\\usuario\\Documents\\archivo.pst" focused onInput={onInput} onSubmit={submitCurrentPath} />
      </box>
    </box>
  );
}

function OptionsPanel({ options, selectedIndex }: { options: ImportOptions; selectedIndex: number }) {
  const rows = [
    `Accion: ${options.action}`,
    `Saltar duplicados: ${options.skipDuplicates ? "Si" : "No"}`,
    `Filtro por anio: ${options.filterYear ?? "Ninguno"}`,
    "Escanear PST y seleccionar carpetas",
    "Volver a ruta PST",
  ];
  return (
    <box title="Opciones de importacion" style={{ border: true, padding: 1, flexDirection: "column" }}>
      {rows.map((row, index) => (
        <text key={row} fg={index === selectedIndex ? "#020617" : "#e2e8f0"} bg={index === selectedIndex ? "#facc15" : undefined}>
          {index === selectedIndex ? " > " : "   "}{row}
        </text>
      ))}
    </box>
  );
}

function FolderPanel({ folders, selectedIndex, selectedFolders, options }: { folders: PstFolder[]; selectedIndex: number; selectedFolders: Set<string>; options: ImportOptions }) {
  const total = selectedFolders.size;
  return (
    <box title="Carpetas detectadas" style={{ border: true, padding: 1, flexDirection: "column" }}>
      <text fg="#a7f3d0">Seleccionadas: {total} · Accion: {options.action} · Duplicados: {options.skipDuplicates ? "saltar" : "importar"}</text>
      {folders.slice(Math.max(0, selectedIndex - 8), selectedIndex + 12).map((folder) => {
        const active = folders[selectedIndex]?.path === folder.path;
        const checked = selectedFolders.has(folder.path);
        const disabled = folder.itemCount <= 0;
        return (
          <text key={folder.path} fg={disabled ? "#64748b" : active ? "#020617" : "#e2e8f0"} bg={active ? "#a3e635" : undefined}>
            {active ? " > " : "   "}[{checked ? "X" : " "}] {folder.path} <span>({folder.itemCount}) {folder.yearSummary ?? ""}</span>
          </text>
        );
      })}
    </box>
  );
}

function StatusPanel({ title, color, lines }: { title: string; color: string; lines: string[] }) {
  return (
    <box title={title} style={{ border: true, padding: 1, flexDirection: "column" }}>
      {lines.map((line, index) => <text key={`${title}-${index}`} fg={color}>{line}</text>)}
    </box>
  );
}

async function main() {
  const renderer = await createCliRenderer({ exitOnCtrlC: true });
  const keymap = createDefaultOpenTuiKeymap(renderer);
  process.on("exit", () => void cleanupScripts());
  createRoot(renderer).render(
    <KeymapProvider keymap={keymap}>
      <App />
    </KeymapProvider>,
  );
}

void main();
