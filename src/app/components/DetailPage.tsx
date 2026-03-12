import type { ReactNode } from "react";
import type { ParsedColumn, ParsedRow } from "../../types";
import { IconChevron } from "../icons";

interface DetailPageProps {
  selectedRow: ParsedRow | null;
  displayColumns: ParsedColumn[];
  hiddenColumns: ParsedColumn[];
  showHiddenFields: boolean;
  onToggleHiddenFields: () => void;
  onOpenAIRunModal: () => void;
  renderDetailField: (column: ParsedColumn, isHidden: boolean) => ReactNode;
}

export function DetailPage({
  selectedRow,
  displayColumns,
  hiddenColumns,
  showHiddenFields,
  onToggleHiddenFields,
  onOpenAIRunModal,
  renderDetailField,
}: DetailPageProps) {
  if (!selectedRow) {
    return (
      <div className="record-list-empty">请先在题目列表页选择一条记录</div>
    );
  }

  return (
    <section className="record-detail standalone-record-detail">
      <div className="record-detail-header">
        <h3>字段详情</h3>
        <span>点击字段左侧勾选框可控制显示/隐藏</span>
      </div>
      <div className="record-detail-ai-toolbar">
        <div className="record-detail-ai-actions">
          <button
            type="button"
            className="btn btn-primary"
            onClick={onOpenAIRunModal}
          >
            运行AI检测
          </button>
        </div>
      </div>
      <div className="detail-fields">
        {displayColumns.map((column) => renderDetailField(column, false))}
        {hiddenColumns.length > 0 ? (
          <div className="hidden-fields-section">
            <button
              type="button"
              className={`hidden-fields-toggle ${showHiddenFields ? "expanded" : ""}`}
              onClick={onToggleHiddenFields}
            >
              <IconChevron />
              <span>{hiddenColumns.length} 个已隐藏字段</span>
            </button>
            {showHiddenFields ? (
              <div className="hidden-fields-list">
                {hiddenColumns.map((column) => renderDetailField(column, true))}
              </div>
            ) : null}
          </div>
        ) : null}
      </div>
    </section>
  );
}
