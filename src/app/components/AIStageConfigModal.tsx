import { useEffect, useState } from "react";
import type { Dispatch, SetStateAction } from "react";
import type {
    AIDetectConfig,
    AIDetectStageKey,
    FileViewState,
    ParsedColumn,
} from "../../types";
import {
    AI_PROVIDER_OPTIONS,
    AI_STAGE_LABELS,
    AI_STAGE_ORDER,
} from "../constants";

interface AIStageConfigModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    aiConfigFormMessage: string;
    draftAIConfig: AIDetectConfig;
    setDraftAIConfig: Dispatch<SetStateAction<AIDetectConfig>>;
    aiSubmitFieldColumns: ParsedColumn[];
    aiConfigSaving: boolean;
    onToggleDraftAISubmitField: (
        stageKey: AIDetectStageKey,
        columnKey: string,
    ) => void;
    onCancel: () => void;
    onSave: () => void;
}

export function AIStageConfigModal({
    isOpen,
    activeFile,
    aiConfigFormMessage,
    draftAIConfig,
    setDraftAIConfig,
    aiSubmitFieldColumns,
    aiConfigSaving,
    onToggleDraftAISubmitField,
    onCancel,
    onSave,
}: AIStageConfigModalProps) {
    const [activeStageKey, setActiveStageKey] =
        useState<AIDetectStageKey>("precheck");

    useEffect(() => {
        if (isOpen) {
            setActiveStageKey("precheck");
        }
    }, [isOpen]);

    if (!isOpen || !activeFile) {
        return null;
    }

    const stageConfig = draftAIConfig.stages[activeStageKey];
    const profileOptions = draftAIConfig.profiles ?? [];
    const activeProfile =
        profileOptions.find((item) => item.name === stageConfig.profileName) ??
        profileOptions[0] ??
        null;
    const getProviderLabel = (
        provider: AIDetectConfig["profiles"][number]["profile"]["provider"],
    ) =>
        AI_PROVIDER_OPTIONS.find((item) => item.value === provider)?.label ??
        "未配置供应商";

    const updateStageConfig = (
        updater: (
            stage: AIDetectConfig["stages"][AIDetectStageKey],
        ) => AIDetectConfig["stages"][AIDetectStageKey],
    ) => {
        setDraftAIConfig((previous) => ({
            ...previous,
            stages: {
                ...previous.stages,
                [activeStageKey]: updater(previous.stages[activeStageKey]),
            },
        }));
    };

    return (
        <div className="column-modal-mask">
            <div className="column-modal ai-config-modal">
                <h3>AI阶段任务配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-config-stage-tabs">
                        {AI_STAGE_ORDER.map((stageKey) => {
                            const label = AI_STAGE_LABELS[stageKey];
                            const isActive = stageKey === activeStageKey;
                            return (
                                <button
                                    key={stageKey}
                                    type="button"
                                    className={`ai-config-stage-tab ${isActive ? "is-active" : ""}`}
                                    onClick={() => setActiveStageKey(stageKey)}
                                >
                                    <span>{label.shortTitle}</span>
                                    <small>{label.title}</small>
                                </button>
                            );
                        })}
                    </div>
                    <div className="ai-config-stage-info">
                        <strong>{AI_STAGE_LABELS[activeStageKey].title}</strong>
                        <span>
                            {AI_STAGE_LABELS[activeStageKey].description}
                        </span>
                    </div>
                    <label className="ai-config-field">
                        <span>选择接口配置</span>
                        <select
                            value={stageConfig.profileName}
                            onChange={(event) =>
                                updateStageConfig((previous) => ({
                                    ...previous,
                                    profileName: event.target.value,
                                }))
                            }
                        >
                            <option value="">请选择接口配置</option>
                            {profileOptions.map((profile) => (
                                <option key={profile.name} value={profile.name}>
                                    {profile.name}
                                </option>
                            ))}
                        </select>
                        {activeProfile ? (
                            <small className="ai-config-hint">
                                {`${getProviderLabel(activeProfile.profile.provider)} · ${activeProfile.profile.model || "未配置模型"}`}
                            </small>
                        ) : (
                            <small className="ai-config-hint">
                                尚未配置接口
                            </small>
                        )}
                    </label>
                    <div className="ai-config-field">
                        <span>结果保存</span>
                        <div className="ai-config-hint">
                            AI 检测结果将直接写入“AI检测结果”区域，不再保存到
                            Excel 字段。
                        </div>
                    </div>
                    <div className="ai-config-section">
                        <div className="ai-config-section-title">
                            提交回答字段（可多选）
                        </div>
                        <div className="ai-config-fields">
                            {aiSubmitFieldColumns.map((column) => {
                                const checked =
                                    stageConfig.submitFieldKeys.includes(
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
                                                onToggleDraftAISubmitField(
                                                    activeStageKey,
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
                        <span>
                            Prompt（推荐使用 <code>{"{{fields_json}}"}</code>
                            ，也支持 <code>{"{{fields_text}}"}</code> /{" "}
                            <code>{"{{image_fields}}"}</code>）
                        </span>
                        <textarea
                            value={stageConfig.prompt}
                            onChange={(event) =>
                                updateStageConfig((previous) => ({
                                    ...previous,
                                    prompt: event.target.value,
                                }))
                            }
                            placeholder="请输入提示词"
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
                        {aiConfigSaving ? "保存中..." : "保存阶段任务配置"}
                    </button>
                </div>
            </div>
        </div>
    );
}
