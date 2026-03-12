import type { ReactNode } from "react";
import type { FileViewState, ParsedColumn, ParsedRow } from "../../types";

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
              const checked = batchSelectedRowIdSet.has(row.rowId);
              const rowNumber = (listPage - 1) * listPageSize + index + 1;
              return (
                <tr
                  key={row.rowId}
                  className={selectedRowId === row.rowId ? "active" : ""}
                  onClick={() => onOpenRowDetail(row.rowId)}
                >
                  <td
                    className="list-table-check-cell"
                    onClick={(event) => event.stopPropagation()}
                  >
                    <input
                      type="checkbox"
                      checked={checked}
                      onChange={() => onToggleBatchRowSelection(row.rowId)}
                    />
                  </td>
                  <td className="list-table-index-cell">{rowNumber}</td>
                  {activeFile.columns.map((column) => {
                    const cell = row.values[column.key];
                    const isImage = cell?.type === "image" && cell.src;
                    return (
                      <td
                        key={`${row.rowId}_${column.key}`}
                        title={getCellTitle(row, column) || undefined}
                      >
                        <div
                          className={
                            isImage ? "list-cell-image" : "list-cell-text"
                          }
                        >
                          {renderListReadonlyCell(row, column)}
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
          <label className="filter-group">
            <span>每页</span>
            <select
              value={listPageSize}
              onChange={(event) => onPageSizeChange(Number(event.target.value))}
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
            onClick={() => onPageChange(Math.min(totalListPages, listPage + 1))}
            disabled={listPage >= totalListPages}
          >
            下一页
          </button>
        </div>
      </div>
    </>
  );
}
