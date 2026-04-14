import { useEffect, useMemo, useState } from "react";
import type { FileViewState, ParsedRow } from "../../types";

function serializeFilters(filters: FileViewState["filterConditions"]): string {
    return JSON.stringify(
        filters.map((item) => ({
            columnKey: item.columnKey,
            value: item.value,
        })),
    );
}

function isParsedRow(value: unknown): value is ParsedRow {
    return (
        !!value &&
        typeof value === "object" &&
        typeof (value as ParsedRow).rowId === "string" &&
        !!(value as ParsedRow).values &&
        typeof (value as ParsedRow).values === "object"
    );
}

export const useListView = ({
    activeFile,
    selectedRowId,
    setSelectedRowId,
    defaultPageSize,
    setErrorMessage,
}: {
    activeFile: FileViewState | null;
    selectedRowId: string | null;
    setSelectedRowId: (value: string | null) => void;
    defaultPageSize: number;
    setErrorMessage: (value: string) => void;
}) => {
    const [listPage, setListPage] = useState(1);
    const [listPageSize, setListPageSize] = useState<number>(defaultPageSize);
    const [batchSelectedRowIds, setBatchSelectedRowIds] = useState<string[]>(
        [],
    );
    const [paginatedRows, setPaginatedRows] = useState<ParsedRow[]>([]);
    const [totalFilteredRows, setTotalFilteredRows] = useState(0);
    const [selectedRow, setSelectedRow] = useState<ParsedRow | null>(null);
    const [previousRowId, setPreviousRowId] = useState<string | null>(null);
    const [nextRowId, setNextRowId] = useState<string | null>(null);
    const [filterOptionsByColumn, setFilterOptionsByColumn] = useState<
        Record<string, string[]>
    >({});
    const [isListLoading, setIsListLoading] = useState(false);
    const [isDetailLoading, setIsDetailLoading] = useState(false);

    const displayColumns = useMemo(() => {
        if (!activeFile) {
            return [];
        }
        return activeFile.columns.filter((column) =>
            activeFile.selectedDisplayColumnKeys.includes(column.key),
        );
    }, [activeFile]);

    const hiddenColumns = useMemo(() => {
        if (!activeFile) {
            return [];
        }
        return activeFile.columns.filter(
            (column) => !activeFile.selectedDisplayColumnKeys.includes(column.key),
        );
    }, [activeFile]);

    const serializedFilters = useMemo(
        () =>
            activeFile
                ? serializeFilters(activeFile.filterConditions)
                : JSON.stringify([]),
        [activeFile],
    );

    const totalListPages = Math.max(
        1,
        Math.ceil(totalFilteredRows / listPageSize) || 1,
    );

    useEffect(() => {
        setListPage(1);
        setBatchSelectedRowIds([]);
    }, [activeFile?.fileId, serializedFilters]);

    useEffect(() => {
        if (listPage > totalListPages) {
            setListPage(totalListPages);
        }
    }, [listPage, totalListPages]);

    useEffect(() => {
        setFilterOptionsByColumn({});
    }, [activeFile?.fileId]);

    useEffect(() => {
        if (!activeFile) {
            setPaginatedRows([]);
            setTotalFilteredRows(0);
            setIsListLoading(false);
            return;
        }

        const controller = new AbortController();
        setIsListLoading(true);

        const loadRows = async () => {
            try {
                const params = new URLSearchParams({
                    page: String(listPage),
                    pageSize: String(listPageSize),
                    filters: serializedFilters,
                });
                const response = await fetch(
                    `/api/files/${encodeURIComponent(activeFile.fileId)}/rows?${params.toString()}`,
                    { signal: controller.signal },
                );
                if (!response.ok) {
                    throw new Error("加载列表失败");
                }
                const payload = (await response.json()) as {
                    rows?: unknown;
                    totalCount?: unknown;
                    page?: unknown;
                };
                const nextRows = Array.isArray(payload.rows)
                    ? payload.rows.filter(isParsedRow)
                    : [];
                const nextTotal =
                    typeof payload.totalCount === "number"
                        ? payload.totalCount
                        : 0;
                const normalizedPage =
                    typeof payload.page === "number" ? payload.page : listPage;
                setPaginatedRows(nextRows);
                setTotalFilteredRows(nextTotal);
                if (normalizedPage !== listPage) {
                    setListPage(normalizedPage);
                }
            } catch (error) {
                if (!controller.signal.aborted) {
                    setErrorMessage(
                        error instanceof Error ? error.message : "加载列表失败",
                    );
                    setPaginatedRows([]);
                    setTotalFilteredRows(0);
                }
            } finally {
                if (!controller.signal.aborted) {
                    setIsListLoading(false);
                }
            }
        };

        void loadRows();
        return () => controller.abort();
    }, [
        activeFile,
        listPage,
        listPageSize,
        serializedFilters,
        setErrorMessage,
    ]);

    useEffect(() => {
        if (!activeFile || !selectedRowId) {
            setSelectedRow(null);
            setPreviousRowId(null);
            setNextRowId(null);
            setIsDetailLoading(false);
            return;
        }

        const controller = new AbortController();
        setIsDetailLoading(true);

        const loadDetail = async () => {
            try {
                const params = new URLSearchParams({
                    filters: serializedFilters,
                });
                const response = await fetch(
                    `/api/files/${encodeURIComponent(activeFile.fileId)}/rows/${encodeURIComponent(selectedRowId)}?${params.toString()}`,
                    { signal: controller.signal },
                );
                if (response.status === 404) {
                    setSelectedRow(null);
                    setPreviousRowId(null);
                    setNextRowId(null);
                    setSelectedRowId(null);
                    return;
                }
                if (!response.ok) {
                    throw new Error("加载详情失败");
                }
                const payload = (await response.json()) as {
                    row?: unknown;
                    previousRowId?: unknown;
                    nextRowId?: unknown;
                };
                setSelectedRow(isParsedRow(payload.row) ? payload.row : null);
                setPreviousRowId(
                    typeof payload.previousRowId === "string"
                        ? payload.previousRowId
                        : null,
                );
                setNextRowId(
                    typeof payload.nextRowId === "string"
                        ? payload.nextRowId
                        : null,
                );
            } catch (error) {
                if (!controller.signal.aborted) {
                    setErrorMessage(
                        error instanceof Error ? error.message : "加载详情失败",
                    );
                    setSelectedRow(null);
                    setPreviousRowId(null);
                    setNextRowId(null);
                }
            } finally {
                if (!controller.signal.aborted) {
                    setIsDetailLoading(false);
                }
            }
        };

        void loadDetail();
        return () => controller.abort();
    }, [
        activeFile,
        selectedRowId,
        serializedFilters,
        setErrorMessage,
        setSelectedRowId,
    ]);

    const batchSelectedRowIdSet = useMemo(
        () => new Set(batchSelectedRowIds),
        [batchSelectedRowIds],
    );

    const onToggleBatchRowSelection = (rowId: string) => {
        setBatchSelectedRowIds((previous) => {
            if (previous.includes(rowId)) {
                return previous.filter((item) => item !== rowId);
            }
            return [...previous, rowId];
        });
    };

    const onSelectAllBatchRows = async () => {
        if (!activeFile) {
            return;
        }
        if (batchSelectedRowIds.length === totalFilteredRows && totalFilteredRows > 0) {
            setBatchSelectedRowIds([]);
            return;
        }
        try {
            const params = new URLSearchParams({ filters: serializedFilters });
            const response = await fetch(
                `/api/files/${encodeURIComponent(activeFile.fileId)}/row-ids?${params.toString()}`,
            );
            if (!response.ok) {
                throw new Error("加载筛选结果失败");
            }
            const payload = (await response.json()) as { rowIds?: unknown };
            const rowIds = Array.isArray(payload.rowIds)
                ? payload.rowIds.filter(
                      (item): item is string => typeof item === "string",
                  )
                : [];
            setBatchSelectedRowIds(rowIds);
        } catch (error) {
            setErrorMessage(
                error instanceof Error
                    ? error.message
                    : "加载筛选结果失败",
            );
        }
    };

    const onSelectCurrentPageBatchRows = () => {
        setBatchSelectedRowIds(paginatedRows.map((row) => row.rowId));
    };

    const onClearBatchRows = () => {
        setBatchSelectedRowIds([]);
    };

    const replaceRowInCaches = (nextRow: ParsedRow) => {
        setPaginatedRows((previous) =>
            previous.map((row) => (row.rowId === nextRow.rowId ? nextRow : row)),
        );
        setSelectedRow((previous) =>
            previous?.rowId === nextRow.rowId ? nextRow : previous,
        );
    };

    const loadFilterOptions = async (columnKey: string): Promise<string[]> => {
        if (!activeFile || columnKey.trim().length === 0) {
            return [];
        }
        const cached = filterOptionsByColumn[columnKey];
        if (cached) {
            return cached;
        }
        const params = new URLSearchParams({ columnKey });
        const response = await fetch(
            `/api/files/${encodeURIComponent(activeFile.fileId)}/filter-options?${params.toString()}`,
        );
        if (!response.ok) {
            throw new Error("加载筛选项失败");
        }
        const payload = (await response.json()) as { options?: unknown };
        const options = Array.isArray(payload.options)
            ? payload.options.filter(
                  (item): item is string => typeof item === "string",
              )
            : [];
        setFilterOptionsByColumn((previous) => ({
            ...previous,
            [columnKey]: options,
        }));
        return options;
    };

    return {
        listPage,
        listPageSize,
        setListPage,
        setListPageSize,
        totalListPages,
        paginatedRows,
        totalFilteredRows,
        displayColumns,
        hiddenColumns,
        selectedRow,
        previousRowId,
        nextRowId,
        batchSelectedRowIds,
        batchSelectedRowIdSet,
        onToggleBatchRowSelection,
        onSelectAllBatchRows,
        onSelectCurrentPageBatchRows,
        onClearBatchRows,
        replaceRowInCaches,
        filterOptionsByColumn,
        loadFilterOptions,
        isListLoading,
        isDetailLoading,
    };
};
