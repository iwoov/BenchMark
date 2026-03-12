import type { MainSection, RouteState, SettingsSection } from "./types";

export function parseHashRoute(hash: string): RouteState {
  const normalized = hash.replace(/^#/, "") || "/list";
  const segments = normalized.split("/").filter(Boolean);
  const section = segments[0];
  const subPath = segments[1];

  if (section === "detail") {
    // Parse row ID from URL: #/detail/{rowId}
    const rowId = subPath && subPath.length > 0 ? decodeURIComponent(subPath) : null;
    return { section: "detail", settingsSection: "fields", rowId };
  }

  if (section === "settings") {
    return {
      section: "settings",
      settingsSection: subPath === "ai" ? "ai" : "fields",
      rowId: null,
    };
  }

  return { section: "list", settingsSection: "fields", rowId: null };
}

export function buildHashRoute(
  section: MainSection,
  settingsSection: SettingsSection,
  rowId?: string | null,
): string {
  if (section === "settings") {
    return `#/settings/${settingsSection}`;
  }

  if (section === "detail" && rowId) {
    return `#/detail/${encodeURIComponent(rowId)}`;
  }

  return `#/${section}`;
}
