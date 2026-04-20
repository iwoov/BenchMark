import type { KeyboardEvent } from "react";

interface ProjectNameDialogProps {
    mode: "create" | "rename" | "add-datasource" | null;
    value: string;
    dataSourceNameValue: string;
    errorMessage: string;
    targetProjectName?: string;
    onChange: (value: string) => void;
    onChangeDataSourceName: (value: string) => void;
    onCancel: () => void;
    onConfirm: () => void;
}

export function ProjectNameDialog({
    mode,
    value,
    dataSourceNameValue,
    errorMessage,
    targetProjectName,
    onChange,
    onChangeDataSourceName,
    onCancel,
    onConfirm,
}: ProjectNameDialogProps) {
    if (!mode) {
        return null;
    }

    const title =
        mode === "create"
            ? "新建项目"
            : mode === "add-datasource"
              ? "添加数据源"
              : "重命名项目";
    const description =
        mode === "create"
            ? "请输入项目名称，后续将以该名称在全站展示。"
            : mode === "add-datasource"
              ? `向项目"${targetProjectName ?? "未命名项目"}"添加新数据源。`
              : `正在修改项目名称：${targetProjectName ?? "未命名项目"}`;

    const showDataSourceName = mode === "create" || mode === "add-datasource";
    const projectNameReadOnly = mode === "add-datasource";

    const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
        if (event.key === "Enter") {
            event.preventDefault();
            onConfirm();
        }
    };

    const confirmLabel =
        mode === "create" || mode === "add-datasource"
            ? "继续选择文件"
            : "保存名称";

    return (
        <div className="column-modal-mask">
            <div className="project-name-dialog">
                <h3>{title}</h3>
                <p>{description}</p>
                <label className="project-name-field">
                    <span>项目名称</span>
                    <input
                        autoFocus={!projectNameReadOnly}
                        type="text"
                        value={value}
                        readOnly={projectNameReadOnly}
                        disabled={projectNameReadOnly}
                        onChange={(event) => onChange(event.target.value)}
                        onKeyDown={handleKeyDown}
                        placeholder="例如：化学/生物/材料多模态评测集"
                    />
                </label>
                {showDataSourceName ? (
                    <label className="project-name-field">
                        <span>数据源名称（可选）</span>
                        <input
                            autoFocus={projectNameReadOnly}
                            type="text"
                            value={dataSourceNameValue}
                            onChange={(event) =>
                                onChangeDataSourceName(event.target.value)
                            }
                            onKeyDown={handleKeyDown}
                            placeholder="例如：version1.1"
                        />
                    </label>
                ) : null}
                {errorMessage ? (
                    <div className="project-name-error">{errorMessage}</div>
                ) : null}
                <div className="project-name-actions">
                    <button type="button" className="btn" onClick={onCancel}>
                        取消
                    </button>
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onConfirm}
                    >
                        {confirmLabel}
                    </button>
                </div>
            </div>
        </div>
    );
}
