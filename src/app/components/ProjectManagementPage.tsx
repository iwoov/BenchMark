import { useState } from "react";
import type { KeyboardEvent } from "react";
import type { FileViewState } from "../../types";
import { IconEdit, IconPlus, IconTrash, IconUpload } from "../icons";

interface ProjectManagementPageProps {
    files: FileViewState[];
    activeFileId: string | null;
    isUploading: boolean;
    onSelectFile: (fileId: string) => void;
    onOpenCreateProjectDialog: () => void;
    onOpenAddDatasourceDialog: (fileId: string) => void;
    onStartMergeUpload: (fileId: string) => void;
    onOpenRenameProjectDialog: (fileId: string) => void;
    onRenameDataSource: (fileId: string, newName: string) => Promise<void>;
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

interface ProjectGroup {
    projectId: string;
    projectName: string;
    dataSources: FileViewState[];
}

function groupFilesByProject(files: FileViewState[]): ProjectGroup[] {
    const groupMap = new Map<string, ProjectGroup>();
    for (const file of files) {
        const gid = file.projectId ?? file.fileId;
        if (!groupMap.has(gid)) {
            groupMap.set(gid, {
                projectId: gid,
                projectName: file.fileName,
                dataSources: [],
            });
        }
        groupMap.get(gid)!.dataSources.push(file);
    }
    return Array.from(groupMap.values());
}

interface DataSourceNameEditorProps {
    fileId: string;
    currentName: string;
    hasMultipleSources: boolean;
    onSave: (fileId: string, newName: string) => Promise<void>;
}

function DataSourceNameEditor({
    fileId,
    currentName,
    hasMultipleSources,
    onSave,
}: DataSourceNameEditorProps) {
    const [editing, setEditing] = useState(false);
    const [draft, setDraft] = useState("");
    const [saving, setSaving] = useState(false);

    const displayName =
        currentName ||
        (hasMultipleSources ? `数据源 ${fileId.slice(0, 6)}` : "默认数据源");

    const handleStartEdit = () => {
        setDraft(currentName);
        setEditing(true);
    };

    const handleCancel = () => {
        setEditing(false);
        setDraft("");
    };

    const handleSave = async () => {
        setSaving(true);
        try {
            await onSave(fileId, draft);
            setEditing(false);
        } finally {
            setSaving(false);
        }
    };

    const handleKeyDown = (event: KeyboardEvent<HTMLInputElement>) => {
        if (event.key === "Enter") {
            event.preventDefault();
            void handleSave();
        }
        if (event.key === "Escape") {
            handleCancel();
        }
    };

    if (editing) {
        return (
            <span className="project-datasource-name-editor">
                <input
                    autoFocus
                    type="text"
                    className="project-datasource-name-input"
                    value={draft}
                    onChange={(e) => setDraft(e.target.value)}
                    onKeyDown={handleKeyDown}
                    placeholder="例如：version1.1"
                    disabled={saving}
                />
                <button
                    type="button"
                    className="btn btn-sm btn-primary"
                    onClick={() => void handleSave()}
                    disabled={saving}
                >
                    {saving ? "保存中..." : "保存"}
                </button>
                <button
                    type="button"
                    className="btn btn-sm"
                    onClick={handleCancel}
                    disabled={saving}
                >
                    取消
                </button>
            </span>
        );
    }

    return (
        <span className="project-datasource-name">
            {displayName}
            <button
                type="button"
                className="btn-icon"
                title="编辑数据源名称"
                onClick={handleStartEdit}
            >
                <IconEdit />
            </button>
        </span>
    );
}

export function ProjectManagementPage({
    files,
    activeFileId,
    isUploading,
    onSelectFile,
    onOpenCreateProjectDialog,
    onOpenAddDatasourceDialog,
    onStartMergeUpload,
    onOpenRenameProjectDialog,
    onRenameDataSource,
    onRemoveFile,
    removingFileId,
}: ProjectManagementPageProps) {
    const projectGroups = groupFilesByProject(files);

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

            {projectGroups.length === 0 ? (
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
                    {projectGroups.map((group) => {
                        const isGroupActive = group.dataSources.some(
                            (ds) => ds.fileId === activeFileId,
                        );
                        const representativeFile = group.dataSources[0];
                        const hasMultipleSources = group.dataSources.length > 1;

                        return (
                            <article
                                key={group.projectId}
                                className={`project-card ${isGroupActive ? "active" : ""}`}
                            >
                                <div className="project-card-head">
                                    <div>
                                        <div className="project-card-title-row">
                                            <h3>{group.projectName}</h3>
                                            {isGroupActive ? (
                                                <span className="project-card-badge">
                                                    当前项目
                                                </span>
                                            ) : null}
                                        </div>
                                    </div>
                                </div>

                                {/* Data sources list */}
                                <div className="project-datasource-list">
                                    {group.dataSources.map((ds) => {
                                        const isActiveDs =
                                            ds.fileId === activeFileId;
                                        const isRemoving =
                                            ds.fileId === removingFileId;
                                        return (
                                            <div
                                                key={ds.fileId}
                                                className={`project-datasource-item ${isActiveDs ? "active" : ""}`}
                                            >
                                                <div className="project-datasource-info">
                                                    <DataSourceNameEditor
                                                        fileId={ds.fileId}
                                                        currentName={
                                                            ds.dataSourceName ??
                                                            ""
                                                        }
                                                        hasMultipleSources={
                                                            hasMultipleSources
                                                        }
                                                        onSave={
                                                            onRenameDataSource
                                                        }
                                                    />
                                                    <span className="project-datasource-meta">
                                                        {ds.sourceFileName
                                                            ? `来源：${ds.sourceFileName}`
                                                            : "来源：未记录"}
                                                        {" · "}
                                                        {formatUpdatedAt(
                                                            ds.updatedAt,
                                                        )}
                                                    </span>
                                                </div>
                                                <div className="project-datasource-metrics">
                                                    <span>
                                                        {ds.rowCount ??
                                                            ds.rows.length}{" "}
                                                        题
                                                    </span>
                                                    <span>
                                                        {ds.columns.length} 字段
                                                    </span>
                                                </div>
                                                <div className="project-datasource-actions">
                                                    <button
                                                        type="button"
                                                        className="btn btn-sm"
                                                        onClick={() =>
                                                            onSelectFile(
                                                                ds.fileId,
                                                            )
                                                        }
                                                        disabled={isRemoving}
                                                    >
                                                        {isActiveDs
                                                            ? "当前"
                                                            : "切换"}
                                                    </button>
                                                    <button
                                                        type="button"
                                                        className="btn btn-sm"
                                                        onClick={() =>
                                                            onStartMergeUpload(
                                                                ds.fileId,
                                                            )
                                                        }
                                                        disabled={
                                                            isUploading ||
                                                            isRemoving
                                                        }
                                                    >
                                                        <IconUpload />
                                                        {isUploading
                                                            ? "导入中..."
                                                            : "继续导入"}
                                                    </button>
                                                    <button
                                                        type="button"
                                                        className="btn btn-sm btn-danger"
                                                        onClick={() =>
                                                            onRemoveFile(
                                                                ds.fileId,
                                                            )
                                                        }
                                                        disabled={
                                                            isUploading ||
                                                            isRemoving
                                                        }
                                                    >
                                                        <IconTrash />
                                                        {isRemoving
                                                            ? "删除中..."
                                                            : "删除"}
                                                    </button>
                                                </div>
                                            </div>
                                        );
                                    })}
                                </div>

                                {/* Project-level actions */}
                                <div className="project-card-actions">
                                    <button
                                        type="button"
                                        className="btn"
                                        onClick={() =>
                                            onOpenAddDatasourceDialog(
                                                representativeFile.fileId,
                                            )
                                        }
                                        disabled={isUploading}
                                    >
                                        <IconPlus />
                                        添加数据源
                                    </button>
                                    <button
                                        type="button"
                                        className="btn"
                                        onClick={() =>
                                            onOpenRenameProjectDialog(
                                                representativeFile.fileId,
                                            )
                                        }
                                        disabled={removingFileId !== null}
                                    >
                                        <IconEdit />
                                        重命名
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
