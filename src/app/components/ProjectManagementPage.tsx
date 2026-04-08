import type { FileViewState } from "../../types";
import { IconEdit, IconPlus, IconTrash, IconUpload } from "../icons";

interface ProjectManagementPageProps {
    files: FileViewState[];
    activeFileId: string | null;
    isUploading: boolean;
    onSelectFile: (fileId: string) => void;
    onOpenCreateProjectDialog: () => void;
    onStartMergeUpload: (fileId: string) => void;
    onOpenRenameProjectDialog: (fileId: string) => void;
    onRemoveFile: (fileId: string) => void;
    removingFileId: string | null;
}

function formatUpdatedAt(value?: string): string {
    if (!value) {
        return "未记录更新时间";
    }
    const date = new Date(value);
    if (Number.isNaN(date.getTime())) {
        return value;
    }
    return new Intl.DateTimeFormat("zh-CN", {
        year: "numeric",
        month: "2-digit",
        day: "2-digit",
        hour: "2-digit",
        minute: "2-digit",
    }).format(date);
}

export function ProjectManagementPage({
    files,
    activeFileId,
    isUploading,
    onSelectFile,
    onOpenCreateProjectDialog,
    onStartMergeUpload,
    onOpenRenameProjectDialog,
    onRemoveFile,
    removingFileId,
}: ProjectManagementPageProps) {
    return (
        <div className="project-page">
            <section className="project-hero">
                <div className="project-hero-copy">
                    <span className="project-eyebrow">项目管理</span>
                    <h2>管理项目名称、导入入口和当前工作项目</h2>
                    <p>
                        项目名称将作为全站显示名称保留；原始 Excel / JSON
                        文件名只作为来源记录。
                    </p>
                </div>
                <div className="project-hero-actions">
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onOpenCreateProjectDialog}
                        disabled={isUploading}
                    >
                        <IconPlus />
                        {isUploading ? "导入中..." : "新建项目并导入"}
                    </button>
                </div>
            </section>

            {files.length === 0 ? (
                <section className="project-empty-state">
                    <h3>还没有项目</h3>
                    <p>先创建一个项目并导入 Excel 或 JSON 文件。</p>
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onOpenCreateProjectDialog}
                        disabled={isUploading}
                    >
                        <IconPlus />
                        创建首个项目
                    </button>
                </section>
            ) : (
                <section className="project-grid">
                    {files.map((file) => {
                        const isActive = file.fileId === activeFileId;
                        const isRemoving = file.fileId === removingFileId;
                        return (
                            <article
                                key={file.fileId}
                                className={`project-card ${isActive ? "active" : ""}`}
                            >
                                <div className="project-card-head">
                                    <div>
                                        <div className="project-card-title-row">
                                            <h3>{file.fileName}</h3>
                                            {isActive ? (
                                                <span className="project-card-badge">
                                                    当前项目
                                                </span>
                                            ) : null}
                                        </div>
                                        <p>
                                            {file.sourceFileName
                                                ? `来源文件：${file.sourceFileName}`
                                                : "来源文件：未记录"}
                                        </p>
                                    </div>
                                </div>

                                <div className="project-card-metrics">
                                    <div>
                                        <span>题目数</span>
                                        <strong>{file.rows.length}</strong>
                                    </div>
                                    <div>
                                        <span>字段数</span>
                                        <strong>{file.columns.length}</strong>
                                    </div>
                                    <div>
                                        <span>最近更新</span>
                                        <strong>
                                            {formatUpdatedAt(file.updatedAt)}
                                        </strong>
                                    </div>
                                </div>

                                <div className="project-card-actions">
                                    <button
                                        type="button"
                                        className="btn"
                                        onClick={() => onSelectFile(file.fileId)}
                                        disabled={isRemoving}
                                    >
                                        切换到该项目
                                    </button>
                                    <button
                                        type="button"
                                        className="btn"
                                        onClick={() =>
                                            onStartMergeUpload(file.fileId)
                                        }
                                        disabled={isUploading || isRemoving}
                                    >
                                        <IconUpload />
                                        {isUploading ? "导入中..." : "继续导入"}
                                    </button>
                                    <button
                                        type="button"
                                        className="btn"
                                        onClick={() =>
                                            onOpenRenameProjectDialog(
                                                file.fileId,
                                            )
                                        }
                                        disabled={isRemoving}
                                    >
                                        <IconEdit />
                                        重命名
                                    </button>
                                    <button
                                        type="button"
                                        className="btn btn-danger"
                                        onClick={() => onRemoveFile(file.fileId)}
                                        disabled={isUploading || isRemoving}
                                    >
                                        <IconTrash />
                                        {isRemoving ? "删除中..." : "删除项目"}
                                    </button>
                                </div>
                            </article>
                        );
                    })}
                </section>
            )}
        </div>
    );
}
