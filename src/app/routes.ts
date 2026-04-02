import type { MainSection, RouteState, SettingsSection } from "./types";

export function parseHashRoute(hash: string): RouteState {
    const normalized = hash.replace(/^#/, "") || "/dashboard";
    const segments = normalized.split("/").filter(Boolean);
    const section = segments[0];
    const subPath = segments[1];

    if (section === "settings") {
        return {
            section: "settings",
            settingsSection:
                subPath === "ai"
                    ? "ai"
                    : subPath === "statistics"
                      ? "statistics"
                      : "fields",
            rowId: null,
        };
    }

    if (section === "dashboard") {
        return {
            section: "dashboard",
            settingsSection: "fields",
            rowId: null,
        };
    }

    return {
        section: "list",
        settingsSection: "fields",
        rowId:
            subPath && subPath.length > 0 ? decodeURIComponent(subPath) : null,
    };
}

export function buildHashRoute(
    section: MainSection,
    settingsSection: SettingsSection,
    rowId?: string | null,
): string {
    if (section === "settings") {
        return `#/settings/${settingsSection}`;
    }
    if (section === "dashboard") {
        return "#/dashboard";
    }
    return rowId ? `#/list/${encodeURIComponent(rowId)}` : "#/list";
}
