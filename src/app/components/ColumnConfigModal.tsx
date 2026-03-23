import type { ParsedFile } from "../../types";

interface ColumnConfigModalProps {
    pendingFile: ParsedFile | null;
    pendingConfigMode: "import" | "edit";
    pendingConfigNotice: string;
    pendingSelectedDisplayKeys: string[];
    pendingEditableColumnKeys: string[];
    onPendingSelectAllDisplayColumns: () => void;
    onPendingClearDisplayColumns: () => void;
    onPendingClearEditableColumns: () => void;
    onTogglePendingDisplayColumn: (columnKey: string) => void;
    onTogglePendingEditableColumn: (columnKey: string) => void;
    onCancelPendingFile: () => void;
    onConfirmPendingFile: () => void;
}

export function ColumnConfigModal({
    pendingFile,
    pendingConfigMode,
    pendingConfigNotice,
    pendingSelectedDisplayKeys,
    pendingEditableColumnKeys,
    onPendingSelectAllDisplayColumns,
    onPendingClearDisplayColumns,
    onPendingClearEditableColumns,
    onTogglePendingDisplayColumn,
    onTogglePendingEditableColumn,
    onCancelPendingFile,
    onConfirmPendingFile,
}: ColumnConfigModalProps) {
    if (!pendingFile) {
        return null;
    }

    return (
        <div className="column-modal-mask">
            <div className="column-modal">
                <h3>
                    {pendingConfigMode === "edit"
                        ? "编辑字段展示/可编辑"
                        : "配置字段展示/可编辑"}
                </h3>
                <p>{pendingFile.fileName}</p>
                {pendingConfigNotice ? (
                    <div className="column-modal-notice">
                        {pendingConfigNotice}
                    </div>
                ) : null}
                <div className="column-modal-actions">
                    <button
                        type="button"
                        className="btn"
                        onClick={onPendingSelectAllDisplayColumns}
                    >
                        全选展示
                    </button>
                    <button
                        type="button"
                        className="btn"
                        onClick={onPendingClearDisplayColumns}
                    >
                        清空展示
                    </button>
                    <button
                        type="button"
                        className="btn"
                        onClick={onPendingClearEditableColumns}
                    >
                        清空可编辑
                    </button>
                </div>
                <div className="column-modal-list">
                    {pendingFile.columns.map((column) => {
                        const checkedDisplay =
                            pendingSelectedDisplayKeys.includes(column.key);
                        const checkedEditable =
                            pendingEditableColumnKeys.includes(column.key);
                        return (
                            <div
                                key={column.key}
                                className={`column-config-row ${checkedEditable ? "editable-column-row" : ""}`}
                            >
                                <span className="column-config-name">
                                    {column.title}
                                </span>
                                <label className="column-config-switch">
                                    <input
                                        type="checkbox"
                                        checked={checkedDisplay}
                                        onChange={() =>
                                            onTogglePendingDisplayColumn(
                                                column.key,
                                            )
                                        }
                                    />
                                    <span>展示</span>
                                </label>
                                <label className="column-config-switch">
                                    <input
                                        type="checkbox"
                                        checked={checkedEditable}
                                        onChange={() =>
                                            onTogglePendingEditableColumn(
                                                column.key,
                                            )
                                        }
                                    />
                                    <span>可编辑</span>
                                </label>
                            </div>
                        );
                    })}
                </div>
                <div className="column-modal-footer">
                    <button
                        type="button"
                        className="btn"
                        onClick={onCancelPendingFile}
                    >
                        取消导入
                    </button>
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onConfirmPendingFile}
                    >
                        {pendingConfigMode === "edit"
                            ? "保存配置"
                            : "确认并保存配置"}
                    </button>
                </div>
            </div>
        </div>
    );
}
