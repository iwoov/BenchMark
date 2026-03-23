import { useEffect, useMemo, useState } from "react";
import type { FileViewState } from "../../types";
import { EMPTY_FILTER_VALUE, NON_EMPTY_FILTER_VALUE } from "../constants";
import { getCellText } from "../file-helpers";

export const useListView = ({
    activeFile,
    selectedRowId,
    setSelectedRowId,
    defaultPageSize,
}: {
    activeFile: FileViewState | null;
    selectedRowId: string | null;
    setSelectedRowId: (value: string | null) => void;
    defaultPageSize: number;
}) => {
    const [listPage, setListPage] = useState(1);
    const [listPageSize, setListPageSize] = useState<number>(defaultPageSize);
    const [batchSelectedRowIds, setBatchSelectedRowIds] = useState<string[]>(
        [],
    );

    const displayColumns = useMemo(() => {
        if (!activeFile) {
            return [];
        }
        return activeFile.columns.filter((column) => {
            return activeFile.selectedDisplayColumnKeys.includes(column.key);
        });
    }, [activeFile]);

    const hiddenColumns = useMemo(() => {
        if (!activeFile) {
            return [];
        }
        return activeFile.columns.filter((column) => {
            return !activeFile.selectedDisplayColumnKeys.includes(column.key);
        });
    }, [activeFile]);

    const visibleRows = useMemo(() => {
        if (!activeFile) {
            return [];
        }

        return activeFile.rows.filter((row) => {
            for (const condition of activeFile.filterConditions) {
                const columnKey = condition.columnKey;
                const filterValue = condition.value.trim();
                if (columnKey.length === 0 || filterValue.length === 0) {
                    continue;
                }
                const value = getCellText(row, columnKey).trim();
                if (filterValue === EMPTY_FILTER_VALUE) {
                    if (value.length !== 0) {
                        return false;
                    }
                    continue;
                }
                if (filterValue === NON_EMPTY_FILTER_VALUE) {
                    if (value.length === 0) {
                        return false;
                    }
                    continue;
                }
                if (value !== filterValue) {
                    return false;
                }
            }
            return true;
        });
    }, [activeFile]);

    const totalListPages = Math.max(
        1,
        Math.ceil(visibleRows.length / listPageSize) || 1,
    );
    const paginatedRows = useMemo(() => {
        const start = (listPage - 1) * listPageSize;
        return visibleRows.slice(start, start + listPageSize);
    }, [visibleRows, listPage, listPageSize]);

    useEffect(() => {
        setListPage(1);
    }, [activeFile?.fileId, activeFile?.filterConditions, listPageSize]);

    useEffect(() => {
        if (listPage > totalListPages) {
            setListPage(totalListPages);
        }
    }, [listPage, totalListPages]);

    useEffect(() => {
        if (!activeFile || visibleRows.length === 0) {
            return;
        }

        if (
            selectedRowId !== null &&
            !visibleRows.some((row) => row.rowId === selectedRowId)
        ) {
            setSelectedRowId(null);
        }
    }, [activeFile, visibleRows, selectedRowId, setSelectedRowId]);

    useEffect(() => {
        if (!activeFile) {
            setBatchSelectedRowIds([]);
            return;
        }

        const visibleIdSet = new Set(visibleRows.map((row) => row.rowId));
        setBatchSelectedRowIds((previous) =>
            previous.filter((rowId) => visibleIdSet.has(rowId)),
        );
    }, [activeFile?.fileId, visibleRows]);

    const selectedRow = useMemo(
        () => visibleRows.find((row) => row.rowId === selectedRowId) ?? null,
        [visibleRows, selectedRowId],
    );

    const activeRowIndex = selectedRow
        ? visibleRows.findIndex((row) => row.rowId === selectedRow.rowId)
        : -1;
    const previousRow =
        activeRowIndex > 0 ? visibleRows[activeRowIndex - 1] : null;
    const nextRow =
        activeRowIndex >= 0 && activeRowIndex < visibleRows.length - 1
            ? visibleRows[activeRowIndex + 1]
            : null;

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

    const onSelectAllBatchRows = () => {
        setBatchSelectedRowIds(visibleRows.map((row) => row.rowId));
    };

    const onClearBatchRows = () => {
        setBatchSelectedRowIds([]);
    };

    const listRangeStart =
        visibleRows.length === 0 ? 0 : (listPage - 1) * listPageSize + 1;
    const listRangeEnd = Math.min(listPage * listPageSize, visibleRows.length);

    return {
        listPage,
        listPageSize,
        setListPage,
        setListPageSize,
        totalListPages,
        paginatedRows,
        visibleRows,
        displayColumns,
        hiddenColumns,
        selectedRow,
        activeRowIndex,
        previousRow,
        nextRow,
        batchSelectedRowIds,
        batchSelectedRowIdSet,
        onToggleBatchRowSelection,
        onSelectAllBatchRows,
        onClearBatchRows,
        listRangeStart,
        listRangeEnd,
    };
};
