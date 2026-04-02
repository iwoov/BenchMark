import { useEffect, useMemo, useState } from "react";
import type { RouteState, MainSection, SettingsSection } from "../types";
import { buildHashRoute, parseHashRoute } from "../routes";

export const getInitialRoute = (): RouteState => {
    if (typeof window !== "undefined") {
        return parseHashRoute(window.location.hash);
    }
    return { section: "dashboard", settingsSection: "fields", rowId: null };
};

type NavigateOptions = { replace?: boolean };

export const useRouteState = ({
    initialRoute,
    onRowIdChange,
}: {
    initialRoute?: RouteState;
    onRowIdChange?: (rowId: string | null) => void;
}) => {
    const seedRoute = useMemo(
        () => initialRoute ?? getInitialRoute(),
        [initialRoute],
    );
    const [activeSection, setActiveSection] = useState<MainSection>(
        seedRoute.section,
    );
    const [activeSettingsSection, setActiveSettingsSection] =
        useState<SettingsSection>(seedRoute.settingsSection);

    const navigateToSection = (
        section: MainSection,
        settingsSection: SettingsSection = activeSettingsSection,
        rowId?: string | null,
        options?: NavigateOptions,
    ) => {
        const nextHash = buildHashRoute(section, settingsSection, rowId);
        if (
            typeof window !== "undefined" &&
            window.location.hash !== nextHash
        ) {
            if (options?.replace) {
                window.history.replaceState(null, "", nextHash);
                setActiveSection(section);
                setActiveSettingsSection(settingsSection);
            } else {
                window.location.hash = nextHash;
            }
            return;
        }

        setActiveSection(section);
        setActiveSettingsSection(settingsSection);
    };

    useEffect(() => {
        if (typeof window === "undefined") {
            return;
        }

        const syncRouteState = () => {
            const nextRoute = parseHashRoute(window.location.hash);
            setActiveSection(nextRoute.section);
            setActiveSettingsSection(nextRoute.settingsSection);
            onRowIdChange?.(
                nextRoute.section === "list" ? (nextRoute.rowId ?? null) : null,
            );
        };

        window.addEventListener("hashchange", syncRouteState);
        syncRouteState();

        return () => {
            window.removeEventListener("hashchange", syncRouteState);
        };
    }, [onRowIdChange]);

    return {
        initialRoute: seedRoute,
        activeSection,
        activeSettingsSection,
        navigateToSection,
        setActiveSection,
        setActiveSettingsSection,
    };
};
