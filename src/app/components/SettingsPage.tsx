import type {
    AIDetectConfig,
    FileViewState,
    NamedAIDetectConfig,
    ParsedColumn,
    StatisticsChartType,
} from "../../types";
import {
    AI_CLEANING_TOOL_LABELS,
    AI_CLEANING_TOOL_ORDER,
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
    onToggleStatisticsField: (fieldKey: string) => void;
    onSetStatisticsChartType: (
        fieldKey: string,
        chartType: StatisticsChartType,
    ) => void;
    onOpenAIStageConfigModal: () => void;
    onOpenAIProfileModal: () => void;
    onOpenAIRouteModal: () => void;
    onOpenAIChatConfigModal: () => void;
    onOpenAICleaningConfigModal: () => void;
}

export function SettingsPage({
    activeSettingsSection,
    activeFile,
    displayColumns,
    aiConfigList,
    aiConfig,
    onOpenActiveFileConfig,
    onToggleStatisticsField,
    onSetStatisticsChartType,
    onOpenAIStageConfigModal,
    onOpenAIProfileModal,
    onOpenAIRouteModal,
    onOpenAIChatConfigModal,
    onOpenAICleaningConfigModal,
}: SettingsPageProps) {
    const visibleDisplayColumns = displayColumns;
    const editableColumns = activeFile.columns.filter((column) =>
        activeFile.selectedEditableColumnKeys.includes(column.key),
    );
    const activeConfig = aiConfigList[0]?.config ?? aiConfig;
    const providers = activeConfig.providers ?? [];
    const routes = activeConfig.routes ?? [];
    const chatConfig = activeConfig.chat;
    const cleaningConfig = activeConfig.cleaning;
    const providerMap = new Map(providers.map((item) => [item.name, item]));
    const routeMap = new Map(routes.map((item) => [item.name, item]));
    const chatRoute = routeMap.get(chatConfig.routeName);
    const statisticsFieldSet = new Set(
        activeFile.statisticsConfig.selectedFieldKeys,
    );

    const getProviderTypeLabel = (apiType: string) =>
        AI_PROVIDER_API_TYPE_OPTIONS.find((item) => item.value === apiType)
            ?.label ?? "-";

    const getChartTypeLabel = (chartType: StatisticsChartType) =>
        chartType === "bar"
            ? "柱状图"
            : chartType === "pie"
              ? "饼图"
              : chartType === "line"
                ? "折线图"
                : "表格";

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

            {activeSettingsSection === "statistics" ? (
                <section className="settings-section">
                    <div className="settings-section-head">
                        <h3>统计设置</h3>
                        <p>
                            为当前数据源选择要展示的统计字段，并指定每个字段的图表类型。
                        </p>
                    </div>
                    <div className="settings-grid">
                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>统计字段</h4>
                                <span>{`已启用 ${activeFile.statisticsConfig.selectedFieldKeys.length} 个`}</span>
                            </div>
                            <div className="settings-stat-grid">
                                {activeFile.columns.map((column) => {
                                    const enabled = statisticsFieldSet.has(
                                        column.key,
                                    );
                                    const chartType =
                                        activeFile.statisticsConfig
                                            .chartTypeByField[column.key] ??
                                        "bar";
                                    return (
                                        <article
                                            key={column.key}
                                            className={`settings-stat-card ${enabled ? "is-enabled" : ""}`}
                                        >
                                            <div className="settings-stat-card-head">
                                                <label className="settings-stat-toggle">
                                                    <input
                                                        type="checkbox"
                                                        checked={enabled}
                                                        onChange={() =>
                                                            onToggleStatisticsField(
                                                                column.key,
                                                            )
                                                        }
                                                    />
                                                    <div>
                                                        <strong>
                                                            {column.title}
                                                        </strong>
                                                        <span>
                                                            {enabled
                                                                ? "已加入统计主页"
                                                                : "未展示"}
                                                        </span>
                                                    </div>
                                                </label>
                                                <span className="settings-tag">
                                                    {getChartTypeLabel(
                                                        chartType,
                                                    )}
                                                </span>
                                            </div>
                                            <div className="settings-stat-controls">
                                                <label className="filter-group">
                                                    <span>图表类型</span>
                                                    <select
                                                        value={chartType}
                                                        onChange={(event) =>
                                                            onSetStatisticsChartType(
                                                                column.key,
                                                                event.target
                                                                    .value as StatisticsChartType,
                                                            )
                                                        }
                                                    >
                                                        <option value="bar">
                                                            柱状图
                                                        </option>
                                                        <option value="pie">
                                                            饼图
                                                        </option>
                                                        <option value="line">
                                                            折线图
                                                        </option>
                                                        <option value="table">
                                                            表格
                                                        </option>
                                                    </select>
                                                </label>
                                            </div>
                                        </article>
                                    );
                                })}
                            </div>
                        </div>
                    </div>
                </section>
            ) : null}

            {activeSettingsSection === "ai" ? (
                <section className="settings-section">
                    <div className="settings-section-head">
                        <h3>AI 设置</h3>
                        <p>
                            统一维护模型提供商、模型路由，以及当前文件的阶段任务绑定。
                        </p>
                    </div>
                    <div className="settings-grid">
                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>阶段任务配置</h4>
                                <span>拆分显示检测任务与聊天任务</span>
                            </div>
                            <div className="settings-task-groups">
                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>AI 检测阶段</h5>
                                            <span>{`共 ${AI_STAGE_ORDER.length} 个阶段`}</span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn btn-primary"
                                            onClick={onOpenAIStageConfigModal}
                                        >
                                            管理阶段任务
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        {AI_STAGE_ORDER.map((stageKey) => {
                                            const stageConfig =
                                                activeConfig.stages[stageKey];
                                            const stageLabel =
                                                AI_STAGE_LABELS[stageKey];
                                            const route = routeMap.get(
                                                stageConfig.routeName,
                                            );
                                            const routeProviders = route
                                                ? route.steps
                                                      .map((step) =>
                                                          providerMap.get(
                                                              step.providerName,
                                                          ),
                                                      )
                                                      .filter(
                                                          (
                                                              item,
                                                          ): item is NonNullable<
                                                              typeof item
                                                          > => Boolean(item),
                                                      )
                                                : [];
                                            return (
                                                <div
                                                    key={stageKey}
                                                    className="settings-stage-item"
                                                >
                                                    <div className="settings-stage-title">
                                                        <strong>
                                                            {stageLabel.title}
                                                        </strong>
                                                        <span className="settings-tag">
                                                            {stageConfig.routeName ||
                                                                "未绑定路由"}
                                                        </span>
                                                        <span className="settings-tag">
                                                            {routeProviders.length >
                                                            0
                                                                ? routeProviders
                                                                      .map(
                                                                          (
                                                                              item,
                                                                          ) =>
                                                                              item.name,
                                                                      )
                                                                      .join(
                                                                          " -> ",
                                                                      )
                                                                : "未配置提供商"}
                                                        </span>
                                                    </div>
                                                    <div className="settings-stage-meta">
                                                        <span>{`提交字段 ${stageConfig.submitFieldKeys.length} 个`}</span>
                                                        <span>{`重试 ${route?.retryCount ?? 0} 次`}</span>
                                                        <span>{`模型 ${route?.model || "-"}`}</span>
                                                    </div>
                                                    <div className="settings-stage-prompt">
                                                        {getPromptPreview(
                                                            stageConfig.prompt,
                                                        )}
                                                    </div>
                                                </div>
                                            );
                                        })}
                                    </div>
                                </div>

                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>题目详情聊天</h5>
                                            <span>
                                                独立于检测阶段的聊天任务
                                            </span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn"
                                            onClick={onOpenAIChatConfigModal}
                                        >
                                            配置聊天任务
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        <div className="settings-stage-item">
                                            <div className="settings-stage-title">
                                                <strong>聊天任务配置</strong>
                                                <span className="settings-tag">
                                                    {chatConfig.routeName ||
                                                        "未绑定路由"}
                                                </span>
                                                <span className="settings-tag">
                                                    {chatRoute
                                                        ? chatRoute.steps
                                                              .map(
                                                                  (step) =>
                                                                      step.providerName,
                                                              )
                                                              .join(" -> ")
                                                        : "未配置提供商"}
                                                </span>
                                            </div>
                                            <div className="settings-stage-meta">
                                                <span>{`默认字段 ${chatConfig.defaultSubmitFieldKeys.length} 个`}</span>
                                                <span>{`模型 ${chatRoute?.model || "-"}`}</span>
                                            </div>
                                            <div className="settings-stage-prompt">
                                                {getPromptPreview(
                                                    chatConfig.prompt,
                                                )}
                                            </div>
                                        </div>
                                    </div>
                                </div>

                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>数据清洗阶段</h5>
                                            <span>
                                                独立于检测与聊天的结构化清洗工具
                                            </span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn"
                                            onClick={
                                                onOpenAICleaningConfigModal
                                            }
                                        >
                                            配置清洗工具
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        {AI_CLEANING_TOOL_ORDER.map(
                                            (toolKey) => {
                                                const toolConfig =
                                                    cleaningConfig[toolKey];
                                                const toolLabel =
                                                    AI_CLEANING_TOOL_LABELS[
                                                        toolKey
                                                    ];
                                                const route = routeMap.get(
                                                    toolConfig.routeName,
                                                );
                                                const providerSummary = route
                                                    ? route.steps
                                                          .map((step) => {
                                                              const provider =
                                                                  providerMap.get(
                                                                      step.providerName,
                                                                  );
                                                              return (
                                                                  provider?.name ??
                                                                  step.providerName
                                                              );
                                                          })
                                                          .join(" -> ")
                                                    : "未配置提供商";
                                                const mappedCount =
                                                    toolConfig.outputMappings.filter(
                                                        (item) =>
                                                            item.targetFieldKey.trim()
                                                                .length > 0,
                                                    ).length;
                                                return (
                                                    <div
                                                        key={toolKey}
                                                        className="settings-stage-item"
                                                    >
                                                        <div className="settings-stage-title">
                                                            <strong>
                                                                {
                                                                    toolLabel.title
                                                                }
                                                            </strong>
                                                            <span className="settings-tag">
                                                                {toolConfig.routeName ||
                                                                    "未绑定路由"}
                                                            </span>
                                                            <span className="settings-tag">
                                                                {
                                                                    providerSummary
                                                                }
                                                            </span>
                                                        </div>
                                                        <div className="settings-stage-meta">
                                                            <span>{`提交字段 ${toolConfig.submitFieldKeys.length} 个`}</span>
                                                            <span>{`输出映射 ${mappedCount}/${toolLabel.outputKeys.length}`}</span>
                                                            <span>
                                                                {toolConfig.autoFillEnabled
                                                                    ? "自动回填开启"
                                                                    : "自动回填关闭"}
                                                            </span>
                                                            <span>{`模型 ${route?.model || "-"}`}</span>
                                                        </div>
                                                        <div className="settings-stage-prompt">
                                                            {getPromptPreview(
                                                                toolConfig.prompt,
                                                            )}
                                                        </div>
                                                    </div>
                                                );
                                            },
                                        )}
                                    </div>
                                </div>
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
                                                    {getProviderTypeLabel(
                                                        provider.apiType,
                                                    )}
                                                </span>
                                                <span className="settings-config-url">
                                                    {provider.apiUrl ||
                                                        "未配置 API URL"}
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
                                                <span>
                                                    {route.model ||
                                                        "未配置模型"}
                                                </span>
                                                <span>{`重试 ${route.retryCount} 次`}</span>
                                                <span className="settings-config-url">
                                                    {route.steps.length > 0
                                                        ? route.steps
                                                              .map(
                                                                  (step) =>
                                                                      step.providerName,
                                                              )
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
