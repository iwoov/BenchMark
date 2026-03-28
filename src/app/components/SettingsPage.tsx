import type {
    AIDetectConfig,
    FileViewState,
    NamedAIDetectConfig,
    ParsedColumn,
} from "../../types";
import {
    AI_PROVIDER_API_TYPE_OPTIONS,
    AI_STAGE_LABELS,
    AI_STAGE_ORDER,
} from "../constants";
import type { SettingsSection } from "../types";

interface SettingsPageProps {
    activeSettingsSection: SettingsSection;
    activeFile: FileViewState;
    displayColumns: ParsedColumn[];
    aiConfigList: NamedAIDetectConfig[];
    aiConfig: AIDetectConfig;
    onOpenActiveFileConfig: () => void;
    onOpenAIStageConfigModal: () => void;
    onOpenAIProfileModal: () => void;
    onOpenAIRouteModal: () => void;
}

export function SettingsPage({
    activeSettingsSection,
    activeFile,
    displayColumns,
    aiConfigList,
    aiConfig,
    onOpenActiveFileConfig,
    onOpenAIStageConfigModal,
    onOpenAIProfileModal,
    onOpenAIRouteModal,
}: SettingsPageProps) {
    const visibleDisplayColumns = displayColumns;
    const editableColumns = activeFile.columns.filter((column) =>
        activeFile.selectedEditableColumnKeys.includes(column.key),
    );
    const activeConfig = aiConfigList[0]?.config ?? aiConfig;
    const providers = activeConfig.providers ?? [];
    const routes = activeConfig.routes ?? [];
    const providerMap = new Map(providers.map((item) => [item.name, item]));
    const routeMap = new Map(routes.map((item) => [item.name, item]));

    const getProviderTypeLabel = (apiType: string) =>
        AI_PROVIDER_API_TYPE_OPTIONS.find((item) => item.value === apiType)
            ?.label ?? "-";

    const getPromptPreview = (prompt: string) => {
        const trimmed = prompt.trim();
        if (!trimmed) {
            return "未配置 Prompt";
        }
        return trimmed.length > 60 ? `${trimmed.slice(0, 60)}…` : trimmed;
    };

    return (
        <div className="settings-layout">
            {activeSettingsSection === "fields" ? (
                <section className="settings-section">
                    <div className="settings-section-head settings-section-head-with-actions">
                        <div className="settings-section-title">
                            <h3>字段展示规则</h3>
                            <p>
                                统一控制详情页哪些字段显示，以及哪些字段允许编辑。
                            </p>
                        </div>
                        <button
                            type="button"
                            className="btn btn-primary"
                            onClick={onOpenActiveFileConfig}
                        >
                            编辑字段
                        </button>
                    </div>
                    <div className="settings-grid">
                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>展示字段</h4>
                                <span>{`已选择 ${visibleDisplayColumns.length} 个`}</span>
                            </div>
                            <div className="settings-pill-list">
                                {visibleDisplayColumns
                                    .slice(0, 12)
                                    .map((column) => (
                                        <span
                                            key={column.key}
                                            className="settings-pill"
                                        >
                                            {column.title}
                                        </span>
                                    ))}
                                {visibleDisplayColumns.length > 12 ? (
                                    <span className="settings-pill">{`还有 ${
                                        visibleDisplayColumns.length - 12
                                    } 个字段`}</span>
                                ) : null}
                            </div>
                        </div>
                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>可编辑字段</h4>
                                <span>{`已选择 ${editableColumns.length} 个`}</span>
                            </div>
                            <div className="settings-pill-list">
                                {editableColumns.length > 0 ? (
                                    editableColumns
                                        .slice(0, 12)
                                        .map((column) => (
                                            <span
                                                key={column.key}
                                                className="settings-pill"
                                            >
                                                {column.title}
                                            </span>
                                        ))
                                ) : (
                                    <span className="settings-empty">
                                        暂无可编辑字段
                                    </span>
                                )}
                                {editableColumns.length > 12 ? (
                                    <span className="settings-pill">{`还有 ${
                                        editableColumns.length - 12
                                    } 个字段`}</span>
                                ) : null}
                            </div>
                        </div>
                    </div>
                </section>
            ) : null}

            {activeSettingsSection === "ai" ? (
                <section className="settings-section">
                    <div className="settings-section-head">
                        <h3>AI 设置</h3>
                        <p>统一维护模型提供商、模型路由，以及当前文件的阶段任务绑定。</p>
                    </div>
                    <div className="settings-grid">
                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>阶段任务</h4>
                                <span>{`共 ${AI_STAGE_ORDER.length} 个阶段`}</span>
                            </div>
                            <div className="settings-stage-list">
                                {AI_STAGE_ORDER.map((stageKey) => {
                                    const stageConfig = activeConfig.stages[stageKey];
                                    const stageLabel = AI_STAGE_LABELS[stageKey];
                                    const route = routeMap.get(stageConfig.routeName);
                                    const routeProviders = route
                                        ? route.steps
                                              .map((step) => providerMap.get(step.providerName))
                                              .filter((item): item is NonNullable<typeof item> => Boolean(item))
                                        : [];
                                    return (
                                        <div
                                            key={stageKey}
                                            className="settings-stage-item"
                                        >
                                            <div className="settings-stage-title">
                                                <strong>{stageLabel.title}</strong>
                                                <span className="settings-tag">
                                                    {stageConfig.routeName || "未绑定路由"}
                                                </span>
                                                <span className="settings-tag">
                                                    {routeProviders.length > 0
                                                        ? routeProviders
                                                              .map((item) => item.name)
                                                              .join(" -> ")
                                                        : "未配置提供商"}
                                                </span>
                                            </div>
                                            <div className="settings-stage-meta">
                                                <span>{`提交字段 ${stageConfig.submitFieldKeys.length} 个`}</span>
                                                <span>{`重试 ${route?.retryCount ?? 0} 次`}</span>
                                                <span>{`模型 ${route?.model || "-"}`}</span>
                                            </div>
                                            <div className="settings-stage-prompt">
                                                {getPromptPreview(stageConfig.prompt)}
                                            </div>
                                        </div>
                                    );
                                })}
                            </div>
                            <div className="settings-section-actions">
                                <button
                                    type="button"
                                    className="btn btn-primary"
                                    onClick={onOpenAIStageConfigModal}
                                >
                                    管理阶段任务
                                </button>
                            </div>
                        </div>

                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>模型提供商</h4>
                                <span>{`共 ${providers.length} 个提供商`}</span>
                            </div>
                            <div className="settings-config-grid">
                                <div className="settings-config-group">
                                    <h5>已配置提供商</h5>
                                    {providers.length > 0 ? (
                                        providers.map((provider) => (
                                            <div
                                                key={provider.name}
                                                className="settings-config-item"
                                            >
                                                <strong>{provider.name}</strong>
                                                <span>
                                                    {getProviderTypeLabel(provider.apiType)}
                                                </span>
                                                <span className="settings-config-url">
                                                    {provider.apiUrl || "未配置 API URL"}
                                                </span>
                                            </div>
                                        ))
                                    ) : (
                                        <span className="settings-empty">
                                            暂无配置
                                        </span>
                                    )}
                                </div>
                            </div>
                            <div className="settings-section-actions">
                                <button
                                    type="button"
                                    className="btn"
                                    onClick={onOpenAIProfileModal}
                                >
                                    管理模型提供商
                                </button>
                            </div>
                        </div>

                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>模型路由</h4>
                                <span>{`共 ${routes.length} 条路由`}</span>
                            </div>
                            <div className="settings-config-grid">
                                <div className="settings-config-group">
                                    <h5>已配置路由</h5>
                                    {routes.length > 0 ? (
                                        routes.map((route) => (
                                            <div
                                                key={route.name}
                                                className="settings-config-item"
                                            >
                                                <strong>{route.name}</strong>
                                                <span>{route.model || "未配置模型"}</span>
                                                <span>{`重试 ${route.retryCount} 次`}</span>
                                                <span className="settings-config-url">
                                                    {route.steps.length > 0
                                                        ? route.steps
                                                              .map((step) => step.providerName)
                                                              .join(" -> ")
                                                        : "未配置回退步骤"}
                                                </span>
                                            </div>
                                        ))
                                    ) : (
                                        <span className="settings-empty">
                                            暂无配置
                                        </span>
                                    )}
                                </div>
                            </div>
                            <div className="settings-section-actions">
                                <button
                                    type="button"
                                    className="btn"
                                    onClick={onOpenAIRouteModal}
                                >
                                    管理模型路由
                                </button>
                            </div>
                        </div>
                    </div>
                </section>
            ) : null}
        </div>
    );
}
