/** Platform values understood by the API; the label is the prefix used in the combined report. */
export type PlatformKey = "desktop" | "ios" | "android" | "ios_native" | "android_native";

export const PLATFORM_LABELS: Record<PlatformKey, string> = {
  desktop: "Desktop Web",
  ios: "iOS Mobile Web",
  android: "Android Mobile Web",
  ios_native: "Native iOS",
  android_native: "Native Android",
};

/** One upload slot per operating system; iOS and Android slots can hold a mobile web or a native app VPAT. */
export type SlotKey = "desktop" | "ios" | "android";
export type Kind = "web" | "native";

export const SLOTS: { key: SlotKey; name: string; nativeCapable: boolean; hint: string }[] = [
  { key: "desktop", name: "Desktop Web", nativeCapable: false, hint: "e.g. …-WEB-Audit-vpat.docx" },
  { key: "ios", name: "iOS", nativeCapable: true, hint: "Mobile web (…-MWEBIOS-…) or native app VPAT" },
  { key: "android", name: "Android", nativeCapable: true, hint: "Mobile web (…-MWEB-ANDROID-…) or native app VPAT" },
];

export function slotPlatform(slot: SlotKey, kind: Kind): PlatformKey {
  return slot !== "desktop" && kind === "native" ? (`${slot}_native` as PlatformKey) : slot;
}

export const MIN_SOURCES = 2;

export interface InspectResult {
  filename: string;
  ok: boolean;
  error?: string;
  productName?: string;
  criteriaCount?: number;
  tables?: { heading: string; count: number }[];
  guessedPlatform?: PlatformKey | null;
  warnings?: string[];
}

export interface Issue {
  severity: "error" | "warning" | "info";
  category: string;
  message: string;
  platform: PlatformKey | null;
  platformLabel: string | null;
  criterion: string | null;
}

export interface Check {
  name: string;
  passed: boolean;
  details: string[];
}

export interface MergeReport {
  ok: boolean;
  errors?: string[];
  productName?: string;
  filename?: string;
  downloadUrl?: string;
  platforms?: PlatformKey[];
  platformLabels?: Partial<Record<PlatformKey, string>>;
  scopeDifferences?: { page: string; presentIn: string[]; missingFrom: string[]; urls: Record<string, string[]> }[];
  sources?: {
    platform: PlatformKey;
    platformLabel: string;
    filename: string;
    productName: string;
    criteriaCount: number;
    tables: { heading: string; count: number }[];
  }[];
  combined?: {
    uniqueCriteria: number;
    inAllSources: number;
    missingFromSome: { key: string; title: string; missing: string[] }[];
    tables: { key: string; count: number }[];
  };
  matrix?: {
    key: string;
    title: string;
    table: string;
    results: Partial<Record<PlatformKey, string | null>>;
  }[];
  issues?: Issue[];
  issueCounts?: Record<"error" | "warning" | "info", number>;
  checks?: Check[];
  outputValid?: boolean;
}

export type StageStatus = "pending" | "running" | "done" | "error";
