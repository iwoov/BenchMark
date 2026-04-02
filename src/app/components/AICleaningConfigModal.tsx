import { useEffect, useState } from "react";
import type { Dispatch, SetStateAction } from "react";
import type {
    AICleaningToolKey,
    AIDetectConfig,
    FileViewState,
    ParsedColumn,
} from "../../types";
import {
    AI_CLEANING_TOOL_LABELS,
    AI_CLEANING_TOOL_ORDER,
    AI_PROVIDER_API_TYPE_OPTIONS,
} from "../constants";

interface AICleaningConfigModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    aiConfigFormMessage: string;
    draftAIConfig: AIDetectConfig;
    setDraftAIConfig: Dispatch<SetStateAction<AIDetectConfig>>;
    aiSubmitFieldColumns: ParsedColumn[];
    aiConfigSaving: boolean;
    onToggleDraftAICleaningSubmitField: (
        toolKey: AICleaningToolKey,
        columnKey: string,
    ) => void;
    onUpdateDraftAICleaningOutputMapping: (
        toolKey: AICleaningToolKey,
        outputKey: string,
        targetFieldKey: string,
    ) => void;
    onCancel: () => void;
    onSave: () => void;
}

export function AICleaningConfigModal({
    isOpen,
    activeFile,
    aiConfigFormMessage,
    draftAIConfig,
    setDraftAIConfig,
    aiSubmitFieldColumns,
    aiConfigSaving,
    onToggleDraftAICleaningSubmitField,
    onUpdateDraftAICleaningOutputMapping,
    onCancel,
    onSave,
}: AICleaningConfigModalProps) {
    const [activeToolKey, setActiveToolKey] =
        useState<AICleaningToolKey>("generate_level3_tags");

    useEffect(() => {
        if (isOpen) {
            setActiveToolKey("generate_level3_tags");
        }
    }, [isOpen]);

    if (!isOpen || !activeFile) {
        return null;
    }

    const toolConfig = draftAIConfig.cleaning[activeToolKey];
    const toolLabel = AI_CLEANING_TOOL_LABELS[activeToolKey];
    const routeOptions = draftAIConfig.routes ?? [];
    const providerMap = new Map(
        draftAIConfig.providers.map((provider) => [provider.name, provider]),
    );
    const activeRoute =
        routeOptions.find((item) => item.name === toolConfig.routeName) ??
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
                <h3>数据清洗配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-config-stage-tabs">
                        {AI_CLEANING_TOOL_ORDER.map((toolKey) => {
                            const label = AI_CLEANING_TOOL_LABELS[toolKey];
                            const isActive = toolKey === activeToolKey;
                            return (
                                <button
                                    key={toolKey}
                                    type="button"
                                    className={`ai-config-stage-tab ${isActive ? "is-active" : ""}`}
                                    onClick={() => setActiveToolKey(toolKey)}
                                >
                                    <span>{label.shortTitle}</span>
                                    <small>{label.title}</small>
                                </button>
                            );
                        })}
                    </div>
                    <div className="ai-config-stage-info">
                        <strong>{toolLabel.title}</strong>
                        <span>{toolLabel.description}</span>
                    </div>
                    <label className="ai-config-field">
                        <span>模型路由</span>
                        <select
                            value={toolConfig.routeName}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    cleaning: {
                                        ...previous.cleaning,
                                        [activeToolKey]: {
                                            ...previous.cleaning[activeToolKey],
                                            routeName: event.target.value,
                                        },
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
                            提交字段（可多选）
                        </div>
                        <div className="ai-config-fields">
                            {aiSubmitFieldColumns.map((column) => {
                                const checked =
                                    toolConfig.submitFieldKeys.includes(
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
                                                onToggleDraftAICleaningSubmitField(
                                                    activeToolKey,
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
                    <div className="ai-config-section">
                        <div className="ai-config-section-title">
                            输出字段映射
                        </div>
                        <label className="ai-config-field">
                            <span>自动回填</span>
                            <label className="ai-config-toggle">
                                <input
                                    type="checkbox"
                                    checked={toolConfig.autoFillEnabled}
                                    onChange={(event) =>
                                        setDraftAIConfig((previous) => ({
                                            ...previous,
                                            cleaning: {
                                                ...previous.cleaning,
                                                [activeToolKey]: {
                                                    ...previous.cleaning[
                                                        activeToolKey
                                                    ],
                                                    autoFillEnabled:
                                                        event.target.checked,
                                                },
                                            },
                                        }))
                                    }
                                />
                                <span>将已映射的输出字段自动覆盖回原始字段</span>
                            </label>
                            <small className="ai-config-hint">
                                默认关闭。关闭时仅保存 AI 原始响应，不回填原始表格字段。
                            </small>
                        </label>
                        <div className="ai-config-mapping-list">
                            {toolLabel.outputKeys.map((outputKey) => {
                                const mapping =
                                    toolConfig.outputMappings.find(
                                        (item) => item.outputKey === outputKey,
                                    ) ?? {
                                        outputKey,
                                        targetFieldKey: "",
                                    };
                                return (
                                    <div
                                        key={outputKey}
                                        className="ai-config-mapping-row"
                                    >
                                        <div className="ai-config-mapping-key">
                                            <strong>{outputKey}</strong>
                                            <span>AI 输出 JSON key</span>
                                        </div>
                                        <label className="ai-config-field">
                                            <span>目标字段</span>
                                            <select
                                                value={mapping.targetFieldKey}
                                                onChange={(event) =>
                                                    onUpdateDraftAICleaningOutputMapping(
                                                        activeToolKey,
                                                        outputKey,
                                                        event.target.value,
                                                    )
                                                }
                                            >
                                                <option value="">
                                                    不映射
                                                </option>
                                                {aiSubmitFieldColumns.map(
                                                    (column) => (
                                                        <option
                                                            key={column.key}
                                                            value={column.key}
                                                        >
                                                            {column.title}
                                                        </option>
                                                    ),
                                                )}
                                            </select>
                                        </label>
                                    </div>
                                );
                            })}
                        </div>
                    </div>
                    <label className="ai-config-field ai-config-prompt-field">
                        <span>
                            Prompt（推荐使用 <code>{"{{fields_json}}"}</code>
                            ，也支持 <code>{"{{fields_text}}"}</code> /{" "}
                            <code>{"{{image_fields}}"}</code>）
                        </span>
                        <textarea
                            value={toolConfig.prompt}
                            onChange={(event) =>
                                setDraftAIConfig((previous) => ({
                                    ...previous,
                                    cleaning: {
                                        ...previous.cleaning,
                                        [activeToolKey]: {
                                            ...previous.cleaning[activeToolKey],
                                            prompt: event.target.value,
                                        },
                                    },
                                }))
                            }
                            placeholder="请输入清洗提示词"
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
                        {aiConfigSaving ? "保存中..." : "保存清洗配置"}
                    </button>
                </div>
            </div>
        </div>
    );
}
