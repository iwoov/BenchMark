import type { KeyboardEvent } from "react";

interface ProjectNameDialogProps {
    mode: "create" | "rename" | null;
    value: string;
    errorMessage: string;
    targetProjectName?: string;
    onChange: (value: string) => void;
    onCancel: () => void;
    onConfirm: () => void;
}

export function ProjectNameDialog({
    mode,
    value,
    errorMessage,
    targetProjectName,
    onChange,
    onCancel,
    onConfirm,
}: ProjectNameDialogProps) {
    if (!mode) {
        return null;
    }

    const title = mode === "create" ? "新建项目" : "重命名项目";
    const description =
        mode === "create"
            ? "请输入项目名称，后续将以该名称在全站展示。"
            : `正在修改项目名称：${targetProjectName ?? "未命名项目"}`;

    const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
        if (event.key === "Enter") {
            event.preventDefault();
            onConfirm();
        }
    };

    return (
        <div className="column-modal-mask">
            <div className="project-name-dialog">
                <h3>{title}</h3>
                <p>{description}</p>
                <label className="project-name-field">
                    <span>项目名称</span>
                    <input
                        autoFocus
                        type="text"
                        value={value}
                        onChange={(event) => onChange(event.target.value)}
                        onKeyDown={handleKeyDown}
                        placeholder="例如：化学/生物/材料多模态评测集"
                    />
                </label>
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
                        {mode === "create" ? "继续选择文件" : "保存名称"}
                    </button>
                </div>
            </div>
        </div>
    );
}
