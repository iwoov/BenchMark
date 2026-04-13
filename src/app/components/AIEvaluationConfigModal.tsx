import { useEffect, useState } from "react";
import type { Dispatch, SetStateAction } from "react";
import type { AIDetectConfig, FileViewState, ParsedColumn } from "../../types";
import {
    AI_PROVIDER_API_TYPE_OPTIONS,
    MAX_AI_EVALUATION_ATTEMPT_COUNT,
    MAX_AI_EVALUATION_MAX_CONCURRENCY,
    MIN_AI_EVALUATION_ATTEMPT_COUNT,
    MIN_AI_EVALUATION_MAX_CONCURRENCY,
} from "../constants";

interface AIEvaluationConfigModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    aiConfigFormMessage: string;
    draftAIConfig: AIDetectConfig;
    setDraftAIConfig: Dispatch<SetStateAction<AIDetectConfig>>;
    aiSubmitFieldColumns: ParsedColumn[];
    aiConfigSaving: boolean;
    onToggleDraftAIEvaluationQuestionField: (
        taskId: string,
        columnKey: string,
    ) => void;
    onToggleDraftAIEvaluationAnswerField: (
        taskId: string,
        columnKey: string,
    ) => void;
    onAddDraftAIEvaluationTask: () => void;
    onRemoveDraftAIEvaluationTask: (taskId: string) => void;
    onCancel: () => void;
    onSave: () => void;
}

export function AIEvaluationConfigModal({
    isOpen,
    activeFile,
    aiConfigFormMessage,
    draftAIConfig,
    setDraftAIConfig,
    aiSubmitFieldColumns,
    aiConfigSaving,
    onToggleDraftAIEvaluationQuestionField,
    onToggleDraftAIEvaluationAnswerField,
    onAddDraftAIEvaluationTask,
    onRemoveDraftAIEvaluationTask,
    onCancel,
    onSave,
}: AIEvaluationConfigModalProps) {
    const [activeTaskId, setActiveTaskId] = useState("");

    useEffect(() => {
        if (!isOpen) {
            return;
        }
        setActiveTaskId((previous) => {
            if (
                previous &&
                draftAIConfig.evaluationTasks.some((task) => task.id === previous)
            ) {
                return previous;
            }
            return draftAIConfig.evaluationTasks[0]?.id ?? "";
        });
    }, [isOpen, draftAIConfig.evaluationTasks]);

    if (!isOpen || !activeFile) {
        return null;
    }

    const activeTask =
        draftAIConfig.evaluationTasks.find((task) => task.id === activeTaskId) ??
        draftAIConfig.evaluationTasks[0];
    if (!activeTask) {
        return null;
    }

    const routeOptions = draftAIConfig.routes ?? [];
    const providerMap = new Map(
        draftAIConfig.providers.map((provider) => [provider.name, provider]),
    );
    const buildProviderSummary = (routeName: string) => {
        const route =
            routeOptions.find((item) => item.name === routeName) ?? null;
        if (!route) {
            return "尚未配置路由";
        }
        return route.steps
            .map((step) => {
                const provider = providerMap.get(step.providerName);
                if (!provider) {
                    return step.providerName;
                }
                const typeLabel =
                    AI_PROVIDER_API_TYPE_OPTIONS.find(
                        (item) => item.value === provider.apiType,
                    )?.label ?? provider.apiType;
                return `${provider.name} (${typeLabel})`;
            })
            .join(" -> ");
    };

    return (
        <div className="column-modal-mask">
            <div className="column-modal ai-config-modal">
                <h3>数据评测配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-profile-actions">
                        <div className="ai-config-stage-info">
                            <strong>多评测任务</strong>
                            <span>
                                每个任务都拥有自己的模型、Prompt、字段和评测次数。详情页可按任务切换展示和运行。
                            </span>
                        </div>
                        <button
                            type="button"
                            className="btn"
                            onClick={onAddDraftAIEvaluationTask}
                        >
                            新增评测任务
                        </button>
                    </div>

                    <div className="ai-config-stage-tabs">
                        {draftAIConfig.evaluationTasks.map((task) => (
                            <button
                                key={task.id}
                                type="button"
                                className={`ai-config-stage-tab ${task.id === activeTask.id ? "is-active" : ""}`}
                                onClick={() => setActiveTaskId(task.id)}
                            >
                                <span>{task.name}</span>
                                <small>
                                    {task.enabled ? "已启用" : "未启用"}
                                </small>
                            </button>
                        ))}
                    </div>

                    <label className="ai-config-field">
                        <span>任务名称</span>
                        <input
                            value={activeTask.name}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    evaluationTasks: previous.evaluationTasks.map(
                                        (task) =>
                                            task.id === activeTask.id
                                                ? {
                                                      ...task,
                                                      name: event.target.value,
                                                  }
                                                : task,
                                    ),
                                }))
                            }
                            placeholder="例如：Gemini 评测"
                        />
                    </label>

                    <div className="ai-profile-actions">
                        <label className="ai-config-toggle">
                            <input
                                type="checkbox"
                                checked={activeTask.enabled}
                                onChange={(event) =>
                                    setDraftAIConfig((previous) => ({
                                        ...previous,
                                        evaluationTasks:
                                            previous.evaluationTasks.map(
                                                (task) =>
                                                    task.id === activeTask.id
                                                        ? {
                                                              ...task,
                                                              enabled:
                                                                  event.target
                                                                      .checked,
                                                          }
                                                        : task,
                                            ),
                                    }))
                                }
                            />
                            <span>启用当前评测任务</span>
                        </label>
                        <button
                            type="button"
                            className="btn"
                            onClick={() => {
                                const nextTasks = draftAIConfig.evaluationTasks;
                                const nextIndex = nextTasks.findIndex(
                                    (task) => task.id === activeTask.id,
                                );
                                const fallbackTask =
                                    nextTasks[nextIndex - 1] ??
                                    nextTasks[nextIndex + 1] ??
                                    null;
                                onRemoveDraftAIEvaluationTask(activeTask.id);
                                if (fallbackTask) {
                                    setActiveTaskId(fallbackTask.id);
                                }
                            }}
                            disabled={draftAIConfig.evaluationTasks.length <= 1}
                        >
                            删除当前任务
                        </button>
                    </div>

                    <label className="ai-config-field">
                        <span>评测次数</span>
                        <input
                            type="number"
                            min={MIN_AI_EVALUATION_ATTEMPT_COUNT}
                            max={MAX_AI_EVALUATION_ATTEMPT_COUNT}
                            value={activeTask.attemptCount}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    evaluationTasks: previous.evaluationTasks.map(
                                        (task) =>
                                            task.id === activeTask.id
                                                ? {
                                                      ...task,
                                                      attemptCount:
                                                          Number.parseInt(
                                                              event.target
                                                                  .value,
                                                              10,
                                                          ) || 0,
                                                  }
                                                : task,
                                    ),
                                }))
                            }
                        />
                    </label>
                    <label className="ai-config-field">
                        <span>最大并发数</span>
                        <input
                            type="number"
                            min={MIN_AI_EVALUATION_MAX_CONCURRENCY}
                            max={MAX_AI_EVALUATION_MAX_CONCURRENCY}
                            value={activeTask.maxConcurrency}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    evaluationTasks: previous.evaluationTasks.map(
                                        (task) =>
                                            task.id === activeTask.id
                                                ? {
                                                      ...task,
                                                      maxConcurrency:
                                                          Number.parseInt(
                                                              event.target
                                                                  .value,
                                                              10,
                                                          ) || 0,
                                                  }
                                                : task,
                                    ),
                                }))
                            }
                        />
                        <small className="ai-config-hint">
                            {`默认 ${5}，允许设置 ${MIN_AI_EVALUATION_MAX_CONCURRENCY} 到 ${MAX_AI_EVALUATION_MAX_CONCURRENCY}。`}
                        </small>
                    </label>

                    <div className="ai-config-section">
                        <div className="ai-config-stage-info">
                            <strong>第一步：题目作答</strong>
                            <span>只基于题目字段独立回答。</span>
                        </div>
                        <label className="ai-config-field">
                            <span>模型路由</span>
                            <select
                                value={activeTask.answerGeneration.routeName}
                                onChange={(event) =>
                                    setDraftAIConfig((previous) => ({
                                        ...previous,
                                        evaluationTasks:
                                            previous.evaluationTasks.map(
                                                (task) =>
                                                    task.id === activeTask.id
                                                        ? {
                                                              ...task,
                                                              answerGeneration:
                                                                  {
                                                                      ...task.answerGeneration,
                                                                      routeName:
                                                                          event
                                                                              .target
                                                                              .value,
                                                                  },
                                                          }
                                                        : task,
                                            ),
                                    }))
                                }
                            >
                                <option value="">请选择模型路由</option>
                                {routeOptions.map((route) => (
                                    <option key={route.name} value={route.name}>
                                        {route.name}
                                    </option>
                                ))}
                            </select>
                            <small className="ai-config-hint">
                                {buildProviderSummary(
                                    activeTask.answerGeneration.routeName,
                                )}
                            </small>
                        </label>
                        <label className="ai-config-field ai-config-prompt-field">
                            <span>第一步 Prompt</span>
                            <textarea
                                value={activeTask.answerGeneration.prompt}
                                onChange={(event) =>
                                    setDraftAIConfig((previous) => ({
                                        ...previous,
                                        evaluationTasks:
                                            previous.evaluationTasks.map(
                                                (task) =>
                                                    task.id === activeTask.id
                                                        ? {
                                                              ...task,
                                                              answerGeneration:
                                                                  {
                                                                      ...task.answerGeneration,
                                                                      prompt:
                                                                          event
                                                                              .target
                                                                              .value,
                                                                  },
                                                          }
                                                        : task,
                                            ),
                                    }))
                                }
                            />
                        </label>
                        <div className="ai-config-section">
                            <div className="ai-config-section-title">
                                题目字段（可多选）
                            </div>
                            <div className="ai-config-fields">
                                {aiSubmitFieldColumns.map((column) => (
                                    <label
                                        key={column.key}
                                        className="ai-config-field-item"
                                    >
                                        <input
                                            type="checkbox"
                                            checked={activeTask.answerGeneration.questionFieldKeys.includes(
                                                column.key,
                                            )}
                                            onChange={() =>
                                                onToggleDraftAIEvaluationQuestionField(
                                                    activeTask.id,
                                                    column.key,
                                                )
                                            }
                                        />
                                        <span>{column.title}</span>
                                    </label>
                                ))}
                            </div>
                        </div>
                    </div>

                    <div className="ai-config-section">
                        <div className="ai-config-stage-info">
                            <strong>第二步：答案判定</strong>
                            <span>系统会自动带上第一步模型回答。</span>
                        </div>
                        <label className="ai-config-field">
                            <span>模型路由</span>
                            <select
                                value={activeTask.answerJudgment.routeName}
                                onChange={(event) =>
                                    setDraftAIConfig((previous) => ({
                                        ...previous,
                                        evaluationTasks:
                                            previous.evaluationTasks.map(
                                                (task) =>
                                                    task.id === activeTask.id
                                                        ? {
                                                              ...task,
                                                              answerJudgment: {
                                                                  ...task.answerJudgment,
                                                                  routeName:
                                                                      event
                                                                          .target
                                                                          .value,
                                                              },
                                                          }
                                                        : task,
                                            ),
                                    }))
                                }
                            >
                                <option value="">请选择模型路由</option>
                                {routeOptions.map((route) => (
                                    <option key={route.name} value={route.name}>
                                        {route.name}
                                    </option>
                                ))}
                            </select>
                            <small className="ai-config-hint">
                                {buildProviderSummary(
                                    activeTask.answerJudgment.routeName,
                                )}
                            </small>
                        </label>
                        <label className="ai-config-field ai-config-prompt-field">
                            <span>第二步 Prompt</span>
                            <textarea
                                value={activeTask.answerJudgment.prompt}
                                onChange={(event) =>
                                    setDraftAIConfig((previous) => ({
                                        ...previous,
                                        evaluationTasks:
                                            previous.evaluationTasks.map(
                                                (task) =>
                                                    task.id === activeTask.id
                                                        ? {
                                                              ...task,
                                                              answerJudgment: {
                                                                  ...task.answerJudgment,
                                                                  prompt:
                                                                      event
                                                                          .target
                                                                          .value,
                                                              },
                                                          }
                                                        : task,
                                            ),
                                    }))
                                }
                            />
                        </label>
                        <div className="ai-config-section">
                            <div className="ai-config-section-title">
                                标准答案字段（可多选）
                            </div>
                            <div className="ai-config-fields">
                                {aiSubmitFieldColumns.map((column) => (
                                    <label
                                        key={column.key}
                                        className="ai-config-field-item"
                                    >
                                        <input
                                            type="checkbox"
                                            checked={activeTask.answerJudgment.answerFieldKeys.includes(
                                                column.key,
                                            )}
                                            onChange={() =>
                                                onToggleDraftAIEvaluationAnswerField(
                                                    activeTask.id,
                                                    column.key,
                                                )
                                            }
                                        />
                                        <span>{column.title}</span>
                                    </label>
                                ))}
                            </div>
                        </div>
                    </div>
                </div>
                <div className="column-modal-footer">
                    <button
                        type="button"
                        className="btn"
                        onClick={onCancel}
                        disabled={aiConfigSaving}
                    >
                        取消
                    </button>
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onSave}
                        disabled={aiConfigSaving}
                    >
                        {aiConfigSaving ? "保存中..." : "保存数据评测配置"}
                    </button>
                </div>
            </div>
        </div>
    );
}
