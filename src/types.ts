export type Store = {
  displayName: string;
  storeId?: string | null;
  filePath?: string | null;
  fileSize?: string | null;
  exchangeStoreType?: number | null;
};

export type StepState = "pending" | "running" | "success" | "error";

export type AppScreen = "preflight" | "stores" | "actions" | "pst-path" | "import-options" | "scan" | "folders" | "done";

export type ActionId = "import-pst" | "backup-pst";

export type PstFolder = {
  path: string;
  itemCount: number;
  yearSummary?: string;
};

export type ImportOptions = {
  action: "Copy" | "Move";
  filterYear?: string;
  filterMonths?: string;
  skipDuplicates: boolean;
};
