import { useEffect, useMemo, useState } from "react";
import type { FileViewState, FilterCondition } from "../../types";
import { EMPTY_FILTER_LABEL, EMPTY_FILTER_VALUE, NON_EMPTY_FILTER_LABEL, NON_EMPTY_FILTER_VALUE } from "../constants";

interface FilterConfigModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    filterOptionsByColumn: Record<string, string[]>;
    loadFilterOptions: (columnKey: string) => Promise<string[]>;
    onClose: () => void;
    onSave: (conditions: FilterCondition[]) => void;
}

function createConditionId(): string {
    return `filter_${Date.now()}_${Math.random().toString(36).slice(2, 8)}`;
}

function getOptionLabel(value: string): string {
    if (value === EMPTY_FILTER_VALUE) {
        return EMPTY_FILTER_LABEL;
    }
    if (value === NON_EMPTY_FILTER_VALUE) {
        return NON_EMPTY_FILTER_LABEL;
    }
    return value;
}

function getInitialCondition(
    activeFile: FileViewState,
    filterOptionsByColumn: Record<string, string[]>,
): FilterCondition | null {
    const firstColumn = activeFile.columns[0];
    if (!firstColumn) {
        return null;
    }
    const options = filterOptionsByColumn[firstColumn.key] ?? [];
    return {
        id: createConditionId(),
        columnKey: firstColumn.key,
        value: options[0] ?? "",
    };
}

export function FilterConfigModal({
    isOpen,
    activeFile,
    filterOptionsByColumn,
    loadFilterOptions,
    onClose,
    onSave,
}: FilterConfigModalProps) {
    const [draftConditions, setDraftConditions] = useState<FilterCondition[]>([]);

    useEffect(() => {
        if (!isOpen || !activeFile) {
            return;
        }
        setDraftConditions(activeFile.filterConditions);
    }, [isOpen, activeFile]);

    useEffect(() => {
        if (!isOpen || !activeFile) {
            return;
        }
        const pendingColumnKeys = new Set<string>();
        draftConditions.forEach((condition) => {
            if (
                condition.columnKey.trim().length > 0 &&
                !filterOptionsByColumn[condition.columnKey]
            ) {
                pendingColumnKeys.add(condition.columnKey);
            }
        });
        if (pendingColumnKeys.size === 0) {
            return;
        }
        pendingColumnKeys.forEach((columnKey) => {
            void loadFilterOptions(columnKey).catch(() => {});
        });
    }, [
        activeFile,
        draftConditions,
        filterOptionsByColumn,
        isOpen,
        loadFilterOptions,
    ]);

    const optionsMap = useMemo(() => {
        if (!activeFile) {
            return new Map<string, string[]>();
        }
        const map = new Map<string, string[]>();
        activeFile.columns.forEach((column) => {
            map.set(column.key, filterOptionsByColumn[column.key] ?? []);
        });
        return map;
    }, [activeFile, filterOptionsByColumn]);

    if (!isOpen || !activeFile) {
        return null;
    }

    const addCondition = () => {
        const next = getInitialCondition(activeFile, filterOptionsByColumn);
        if (next) {
            setDraftConditions((previous) => [...previous, next]);
            return;
        }
        const firstColumn = activeFile.columns[0];
        if (!firstColumn) {
            return;
        }
        setDraftConditions((previous) => [
            ...previous,
            {
                id: createConditionId(),
                columnKey: firstColumn.key,
                value: "",
            },
        ]);
        void loadFilterOptions(firstColumn.key).catch(() => {});
    };

    const updateCondition = (
        conditionId: string,
        updater: (condition: FilterCondition) => FilterCondition,
    ) => {
        setDraftConditions((previous) =>
            previous.map((condition) =>
                condition.id === conditionId ? updater(condition) : condition,
            ),
        );
    };

    const removeCondition = (conditionId: string) => {
        setDraftConditions((previous) =>
            previous.filter((condition) => condition.id !== conditionId),
        );
    };

    const clearConditions = () => {
        setDraftConditions([]);
    };

    const handleSave = () => {
        const validColumnKeys = new Set(activeFile.columns.map((column) => column.key));
        const normalized = draftConditions.filter(
            (condition) =>
                validColumnKeys.has(condition.columnKey) &&
                condition.value.trim().length > 0,
        );
        onSave(normalized);
        onClose();
    };

    return (
        <div className="column-modal-mask">
            <div className="column-modal filter-config-modal">
                <h3>筛选条件</h3>
                <p>{activeFile.fileName}</p>
                <div className="column-modal-actions">
                    <button type="button" className="btn btn-primary" onClick={addCondition}>
                        添加条件
                    </button>
                    <button type="button" className="btn" onClick={clearConditions}>
                        清空条件
                    </button>
                </div>
                <div className="column-modal-list">
                    {draftConditions.length > 0 ? (
                        draftConditions.map((condition, index) => {
                            const valueOptions =
                                optionsMap.get(condition.columnKey) ?? [];
                            return (
                                <div key={condition.id} className="filter-condition-row">
                                    <span className="filter-condition-index">{index + 1}</span>
                                    <label className="filter-condition-field">
                                        <span>字段</span>
                                        <select
                                            value={condition.columnKey}
                                            onChange={(event) => {
                                                const nextColumnKey = event.target.value;
                                                const nextOptions =
                                                    optionsMap.get(nextColumnKey) ?? [];
                                                if (!filterOptionsByColumn[nextColumnKey]) {
                                                    void loadFilterOptions(nextColumnKey).catch(
                                                        () => {},
                                                    );
                                                }
                                                updateCondition(condition.id, (previous) => ({
                                                    ...previous,
                                                    columnKey: nextColumnKey,
                                                    value:
                                                        nextOptions.includes(previous.value)
                                                            ? previous.value
                                                            : (nextOptions[0] ?? ""),
                                                }));
                                            }}
                                        >
                                            {activeFile.columns.map((column) => (
                                                <option key={column.key} value={column.key}>
                                                    {column.title}
                                                </option>
                                            ))}
                                        </select>
                                    </label>
                                    <label className="filter-condition-field">
                                        <span>条件</span>
                                        <select
                                            value={condition.value}
                                            onChange={(event) =>
                                                updateCondition(condition.id, (previous) => ({
                                                    ...previous,
                                                    value: event.target.value,
                                                }))
                                            }
                                        >
                                            {valueOptions.map((value) => (
                                                <option key={value} value={value}>
                                                    {getOptionLabel(value)}
                                                </option>
                                            ))}
                                        </select>
                                    </label>
                                    <button
                                        type="button"
                                        className="btn"
                                        onClick={() => removeCondition(condition.id)}
                                    >
                                        删除
                                    </button>
                                </div>
                            );
                        })
                    ) : (
                        <div className="settings-empty">暂无筛选条件</div>
                    )}
                </div>
                <div className="column-modal-footer">
                    <button type="button" className="btn" onClick={onClose}>
                        取消
                    </button>
                    <button type="button" className="btn btn-primary" onClick={handleSave}>
                        应用筛选
                    </button>
                </div>
            </div>
        </div>
    );
}
