import type { CSSProperties, ReactNode } from "react";
import type {
    AIDetectRunKey,
    AIDetectStageKey,
    FileViewState,
    ParsedColumn,
    ParsedRow,
} from "../../types";
import { AI_STAGE_ORDER } from "../constants";

interface ListPageProps {
    activeFile: FileViewState;
    visibleRows: ParsedRow[];
    paginatedRows: ParsedRow[];
    listPage: number;
    listPageSize: number;
    totalListPages: number;
    listPageSizeOptions: readonly number[];
    batchSelectedRowIdSet: Set<string>;
    selectedRowId: string | null;
    rowStreamProgress?: Record<
        string,
        Partial<Record<AIDetectStageKey, number>>
    >;
    isAIBatchRunning: boolean;
    activeAIRunKey: AIDetectRunKey;
    rowBatchStatuses?: Record<string, "success" | "failed">;
    onToggleBatchRowSelection: (rowId: string) => void;
    onOpenRowDetail: (rowId: string) => void;
    onPageChange: (nextPage: number) => void;
    onPageSizeChange: (nextSize: number) => void;
    getCellTitle: (row: ParsedRow, column: ParsedColumn) => string;
    renderListReadonlyCell: (row: ParsedRow, column: ParsedColumn) => ReactNode;
}

export function ListPage({
    activeFile,
    visibleRows,
    paginatedRows,
    listPage,
    listPageSize,
    totalListPages,
    listPageSizeOptions,
    batchSelectedRowIdSet,
    selectedRowId,
    rowStreamProgress,
    isAIBatchRunning,
    activeAIRunKey,
    rowBatchStatuses,
    onToggleBatchRowSelection,
    onOpenRowDetail,
    onPageChange,
    onPageSizeChange,
    getCellTitle,
    renderListReadonlyCell,
}: ListPageProps) {
    if (visibleRows.length === 0) {
        return <div className="record-list-empty">当前筛选条件下无数据</div>;
    }

    return (
        <>
            <div className="list-table-wrap">
                <table className="list-table">
                    <thead>
                        <tr>
                            <th className="list-table-check-col">选择</th>
                            <th className="list-table-index-col">序号</th>
                            {activeFile.columns.map((column) => (
                                <th key={column.key}>{column.title}</th>
                            ))}
                        </tr>
                    </thead>
                    <tbody>
                        {paginatedRows.map((row, index) => {
                            const checked = batchSelectedRowIdSet.has(
                                row.rowId,
                            );
                            const rowNumber =
                                (listPage - 1) * listPageSize + index + 1;
                            const rowStatus = rowBatchStatuses?.[row.rowId];
                            const completedStages = AI_STAGE_ORDER.reduce(
                                (count, stageKey) => {
                                    const value = row.aiResults?.[stageKey];
                                    return typeof value === "string" &&
                                        value.trim().length > 0
                                        ? count + 1
                                        : count;
                                },
                                0,
                            );
                            let progressPercent = Math.round(
                                (completedStages / AI_STAGE_ORDER.length) * 100,
                            );
                            const streamProgress =
                                rowStreamProgress?.[row.rowId];
                            if (isAIBatchRunning && streamProgress) {
                                if (activeAIRunKey === "all") {
                                    const sum = AI_STAGE_ORDER.reduce(
                                        (acc, stageKey) =>
                                            acc +
                                            (streamProgress[stageKey] ?? 0),
                                        0,
                                    );
                                    const hasAny = AI_STAGE_ORDER.some(
                                        (stageKey) =>
                                            (streamProgress[stageKey] ?? 0) > 0,
                                    );
                                    if (hasAny) {
                                        progressPercent = Math.round(
                                            sum / AI_STAGE_ORDER.length,
                                        );
                                    }
                                } else if (
                                    AI_STAGE_ORDER.includes(
                                        activeAIRunKey as AIDetectStageKey,
                                    )
                                ) {
                                    const value =
                                        streamProgress[
                                            activeAIRunKey as AIDetectStageKey
                                        ];
                                    if (typeof value === "number") {
                                        progressPercent = Math.round(value);
                                    }
                                }
                            }
                            const rowClassName = [
                                selectedRowId === row.rowId ? "active" : "",
                                rowStatus === "failed" ? "ai-row-failed" : "",
                            ]
                                .filter(Boolean)
                                .join(" ");
                            return (
                                <tr
                                    key={row.rowId}
                                    className={rowClassName}
                                    style={
                                        {
                                            "--ai-progress": `${progressPercent}%`,
                                            "--ai-progress-color":
                                                rowStatus === "failed"
                                                    ? "var(--warning)"
                                                    : "var(--accent)",
                                        } as CSSProperties
                                    }
                                    onClick={() => onOpenRowDetail(row.rowId)}
                                >
                                    <td
                                        className="list-table-check-cell"
                                        onClick={(event) =>
                                            event.stopPropagation()
                                        }
                                    >
                                        <input
                                            type="checkbox"
                                            checked={checked}
                                            onChange={() =>
                                                onToggleBatchRowSelection(
                                                    row.rowId,
                                                )
                                            }
                                        />
                                    </td>
                                    <td className="list-table-index-cell">
                                        {rowNumber}
                                    </td>
                                    {activeFile.columns.map((column) => {
                                        const cell = row.values[column.key];
                                        const isImage =
                                            cell?.type === "image" && cell.src;
                                        return (
                                            <td
                                                key={`${row.rowId}_${column.key}`}
                                                title={
                                                    getCellTitle(row, column) ||
                                                    undefined
                                                }
                                            >
                                                <div
                                                    className={
                                                        isImage
                                                            ? "list-cell-image"
                                                            : "list-cell-text"
                                                    }
                                                >
                                                    {renderListReadonlyCell(
                                                        row,
                                                        column,
                                                    )}
                                                </div>
                                            </td>
                                        );
                                    })}
                                </tr>
                            );
                        })}
                    </tbody>
                </table>
            </div>
            <div className="list-pagination">
                <div className="list-pagination-meta">
                    <span>{`当前显示 ${visibleRows.length} / ${activeFile.rows.length} 条`}</span>
                    <label className="filter-group">
                        <span>每页</span>
                        <select
                            value={listPageSize}
                            onChange={(event) =>
                                onPageSizeChange(Number(event.target.value))
                            }
                        >
                            {listPageSizeOptions.map((size) => (
                                <option key={size} value={size}>
                                    {size} 条
                                </option>
                            ))}
                        </select>
                    </label>
                    <span>
                        第 {listPage} / {totalListPages} 页
                    </span>
                </div>
                <div className="list-pagination-actions">
                    <button
                        type="button"
                        className="btn"
                        onClick={() => onPageChange(Math.max(1, listPage - 1))}
                        disabled={listPage <= 1}
                    >
                        上一页
                    </button>
                    <button
                        type="button"
                        className="btn"
                        onClick={() =>
                            onPageChange(Math.min(totalListPages, listPage + 1))
                        }
                        disabled={listPage >= totalListPages}
                    >
                        下一页
                    </button>
                </div>
            </div>
        </>
    );
}
