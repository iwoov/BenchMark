import type { MainSection, RouteState, SettingsSection } from "./types";

export function parseHashRoute(hash: string): RouteState {
  const normalized = hash.replace(/^#/, "") || "/list";
  const segments = normalized.split("/").filter(Boolean);
  const section = segments[0];
  const settingsSection = segments[1];

  if (section === "detail") {
    return { section: "detail", settingsSection: "fields" };
  }

  if (section === "settings") {
    return {
      section: "settings",
      settingsSection: settingsSection === "ai" ? "ai" : "fields",
    };
  }

  return { section: "list", settingsSection: "fields" };
}

export function buildHashRoute(
  section: MainSection,
  settingsSection: SettingsSection,
): string {
  if (section === "settings") {
    return `#/settings/${settingsSection}`;
  }

  return `#/${section}`;
}
