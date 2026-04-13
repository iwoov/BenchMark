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
    onOpenAIEvaluationConfigModal: () => void;
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
    onOpenAIEvaluationConfigModal,
}: SettingsPageProps) {
    const visibleDisplayColumns = displayColumns;
    const editableColumns = activeFile.columns.filter((column) =>
        activeFile.selectedEditableColumnKeys.includes(column.key),
    );
    const activeConfig = aiConfigList[0]?.config ?? aiConfig;
    const providers = activeConfig.providers ?? [];
    const routes = activeConfig.routes ?? [];
    const chatConfig = activeConfig.chat;
    const evaluationTasks = activeConfig.evaluationTasks;
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

    const getRouteProviderSummary = (routeName: string) => {
        const route = routeMap.get(routeName);
        if (!route) {
            return "未配置提供商";
        }
        const names = route.steps
            .map((step) => providerMap.get(step.providerName)?.name ?? step.providerName)
            .filter((name) => name.trim().length > 0);
        return names.length > 0 ? names.join(" -> ") : "未配置提供商";
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
                            按通用配置、数据质检、数据清洗、数据评测四个部分维护当前文件的 AI 流程。
                        </p>
                    </div>
                    <div className="settings-grid">
                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>通用配置</h4>
                                <span>维护共享模型资源与详情页聊天能力</span>
                            </div>
                            <div className="settings-task-groups">
                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>题目详情聊天</h5>
                                            <span>独立于批量质检与清洗的对话任务</span>
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
                                                    {getRouteProviderSummary(
                                                        chatConfig.routeName,
                                                    )}
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
                                            <h5>模型提供商</h5>
                                            <span>{`共 ${providers.length} 个提供商`}</span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn"
                                            onClick={onOpenAIProfileModal}
                                        >
                                            管理模型提供商
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        {providers.length > 0 ? (
                                            providers.map((provider) => (
                                                <div
                                                    key={provider.name}
                                                    className="settings-stage-item"
                                                >
                                                    <div className="settings-stage-title">
                                                        <strong>{provider.name}</strong>
                                                        <span className="settings-tag">
                                                            {getProviderTypeLabel(
                                                                provider.apiType,
                                                            )}
                                                        </span>
                                                    </div>
                                                    <div className="settings-stage-prompt">
                                                        {provider.apiUrl ||
                                                            "未配置 API URL"}
                                                    </div>
                                                </div>
                                            ))
                                        ) : (
                                            <span className="settings-empty">
                                                暂无配置
                                            </span>
                                        )}
                                    </div>
                                </div>

                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>模型路由</h5>
                                            <span>{`共 ${routes.length} 条路由`}</span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn"
                                            onClick={onOpenAIRouteModal}
                                        >
                                            管理模型路由
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        {routes.length > 0 ? (
                                            routes.map((route) => (
                                                <div
                                                    key={route.name}
                                                    className="settings-stage-item"
                                                >
                                                    <div className="settings-stage-title">
                                                        <strong>{route.name}</strong>
                                                        <span className="settings-tag">
                                                            {route.model || "未配置模型"}
                                                        </span>
                                                    </div>
                                                    <div className="settings-stage-meta">
                                                        <span>{`重试 ${route.retryCount} 次`}</span>
                                                        <span>
                                                            {getRouteProviderSummary(
                                                                route.name,
                                                            )}
                                                        </span>
                                                    </div>
                                                </div>
                                            ))
                                        ) : (
                                            <span className="settings-empty">
                                                暂无配置
                                            </span>
                                        )}
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>数据质检</h4>
                                <span>{`共 ${AI_STAGE_ORDER.length} 个检测阶段`}</span>
                            </div>
                            <div className="settings-task-groups">
                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>四阶段质检流程</h5>
                                            <span>按执行顺序维护路由、提交字段和 Prompt</span>
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
                                                            {getRouteProviderSummary(
                                                                stageConfig.routeName,
                                                            )}
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
                            </div>
                        </div>

                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>数据清洗</h4>
                                <span>{`共 ${AI_CLEANING_TOOL_ORDER.length} 个清洗工具`}</span>
                            </div>
                            <div className="settings-task-groups">
                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>结构化清洗工具</h5>
                                            <span>独立维护提交字段、输出映射和自动回填策略</span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn"
                                            onClick={onOpenAICleaningConfigModal}
                                        >
                                            配置清洗工具
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        {AI_CLEANING_TOOL_ORDER.map((toolKey) => {
                                            const toolConfig =
                                                cleaningConfig[toolKey];
                                            const toolLabel =
                                                AI_CLEANING_TOOL_LABELS[toolKey];
                                            const route = routeMap.get(
                                                toolConfig.routeName,
                                            );
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
                                                            {toolLabel.title}
                                                        </strong>
                                                        <span className="settings-tag">
                                                            {toolConfig.routeName ||
                                                                "未绑定路由"}
                                                        </span>
                                                        <span className="settings-tag">
                                                            {getRouteProviderSummary(
                                                                toolConfig.routeName,
                                                            )}
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
                                        })}
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div className="settings-subsection">
                            <div className="settings-subsection-head">
                                <h4>数据评测</h4>
                                <span>独立维护两步评测链路、字段来源与 Prompt</span>
                            </div>
                            <div className="settings-task-groups">
                                <div className="settings-task-group">
                                    <div className="settings-task-group-head">
                                        <div>
                                            <h5>两步评测链路</h5>
                                            <span>先让模型回答题目，再结合标准答案判断是否正确</span>
                                        </div>
                                        <button
                                            type="button"
                                            className="btn btn-primary"
                                            onClick={onOpenAIEvaluationConfigModal}
                                        >
                                            配置数据评测
                                        </button>
                                    </div>
                                    <div className="settings-stage-list">
                                        {evaluationTasks.map((task) => (
                                            <div
                                                key={task.id}
                                                className="settings-stage-item"
                                            >
                                                <div className="settings-stage-title">
                                                    <strong>{task.name}</strong>
                                                    <span className="settings-tag">
                                                        {task.enabled
                                                            ? "已启用"
                                                            : "未启用"}
                                                    </span>
                                                    <span className="settings-tag">
                                                        {`评测次数 ${task.attemptCount} 次`}
                                                    </span>
                                                </div>
                                                <div className="settings-stage-meta">
                                                    <span>{`作答模型 ${routeMap.get(task.answerGeneration.routeName)?.model || "-"}`}</span>
                                                    <span>{`判定模型 ${routeMap.get(task.answerJudgment.routeName)?.model || "-"}`}</span>
                                                    <span>{`题目字段 ${task.answerGeneration.questionFieldKeys.length} 个`}</span>
                                                    <span>{`答案字段 ${task.answerJudgment.answerFieldKeys.length} 个`}</span>
                                                </div>
                                                <div className="settings-stage-prompt">
                                                    {`作答：${getPromptPreview(
                                                        task.answerGeneration.prompt,
                                                    )} ｜ 判定：${getPromptPreview(
                                                        task.answerJudgment.prompt,
                                                    )}`}
                                                </div>
                                            </div>
                                        ))}
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </section>
            ) : null}
        </div>
    );
}
