import type { Dispatch, SetStateAction } from "react";
import type {
    AIDetectConfig,
    FileViewState,
    ParsedColumn,
} from "../../types";
import { AI_PROVIDER_API_TYPE_OPTIONS } from "../constants";

interface AIChatConfigModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    aiConfigFormMessage: string;
    draftAIConfig: AIDetectConfig;
    setDraftAIConfig: Dispatch<SetStateAction<AIDetectConfig>>;
    aiSubmitFieldColumns: ParsedColumn[];
    aiConfigSaving: boolean;
    onToggleDraftAIChatSubmitField: (columnKey: string) => void;
    onCancel: () => void;
    onSave: () => void;
}

export function AIChatConfigModal({
    isOpen,
    activeFile,
    aiConfigFormMessage,
    draftAIConfig,
    setDraftAIConfig,
    aiSubmitFieldColumns,
    aiConfigSaving,
    onToggleDraftAIChatSubmitField,
    onCancel,
    onSave,
}: AIChatConfigModalProps) {
    if (!isOpen || !activeFile) {
        return null;
    }

    const routeOptions = draftAIConfig.routes ?? [];
    const providerMap = new Map(
        draftAIConfig.providers.map((provider) => [provider.name, provider]),
    );
    const activeRoute =
        routeOptions.find((item) => item.name === draftAIConfig.chat.routeName) ??
        routeOptions[0] ??
        null;

    const providerSummary = activeRoute
        ? activeRoute.steps
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
              .join(" -> ")
        : "尚未配置路由";

    return (
        <div className="column-modal-mask">
            <div className="column-modal ai-config-modal">
                <h3>AI 聊天配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-config-stage-info">
                        <strong>题目详情聊天栏</strong>
                        <span>
                            配置聊天默认模型、系统提示词，以及每轮必发的上下文字段。
                        </span>
                    </div>
                    <label className="ai-config-field">
                        <span>默认模型路由</span>
                        <select
                            value={draftAIConfig.chat.routeName}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    chat: {
                                        ...previous.chat,
                                        routeName: event.target.value,
                                    },
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
                        <small className="ai-config-hint">{providerSummary}</small>
                    </label>
                    <div className="ai-config-section">
                        <div className="ai-config-section-title">
                            聊天默认发送字段（可多选）
                        </div>
                        <div className="ai-config-fields">
                            {aiSubmitFieldColumns.map((column) => {
                                const checked =
                                    draftAIConfig.chat.defaultSubmitFieldKeys.includes(
                                        column.key,
                                    );
                                return (
                                    <label
                                        key={column.key}
                                        className="ai-config-field-item"
                                    >
                                        <input
                                            type="checkbox"
                                            checked={checked}
                                            onChange={() =>
                                                onToggleDraftAIChatSubmitField(
                                                    column.key,
                                                )
                                            }
                                        />
                                        <span>{column.title}</span>
                                    </label>
                                );
                            })}
                        </div>
                    </div>
                    <label className="ai-config-field ai-config-prompt-field">
                        <span>聊天系统 Prompt</span>
                        <textarea
                            value={draftAIConfig.chat.prompt}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    chat: {
                                        ...previous.chat,
                                        prompt: event.target.value,
                                    },
                                }))
                            }
                            placeholder="请输入聊天提示词"
                        />
                    </label>
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
                        {aiConfigSaving ? "保存中..." : "保存聊天配置"}
                    </button>
                </div>
            </div>
        </div>
    );
}
