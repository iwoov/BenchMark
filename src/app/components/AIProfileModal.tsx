import type { Dispatch, SetStateAction } from "react";
import type {
    AIDetectConfig,
    AIDetectStageConfigMap,
    FileViewState,
} from "../../types";
import {
    AI_PROVIDER_OPTIONS,
    AI_REASONING_EFFORT_OPTIONS,
    AI_STAGE_ORDER,
    DEFAULT_AI_PROFILE,
    MAX_AI_RETRY_COUNT,
    MIN_AI_RETRY_COUNT,
} from "../constants";
import {
    cloneAIDetectProfile,
    getDefaultAIUrl,
    isGeminiProvider,
    normalizeAIRetryCount,
} from "../ai-helpers";

const MODEL_PROVIDER_OPTIONS = ["openai", "google", "anthropic"] as const;

const MODEL_OPTIONS_BY_PROVIDER: Record<string, string[]> = {
    openai: ["gpt-5.4-2026-03-05", "gpt-5.2-pro-2025-12-11", "gpt-5-pro"],
    google: [
        "gemini-3.1-pro-preview",
        "gemini-3-pro-preview",
        "gemini-3.1-flash-lite-preview",
    ],
    anthropic: ["claude-sonnet-4-6-20260217", "claude-opus-4-6-20260205"],
};

const MODEL_PROVIDER_PREFIXES = new Set([
    "openai",
    "google",
    "anthropic",
    "gemini",
    "vertex",
    "idealab",
]);

const getDefaultModelProvider = (
    provider: AIDetectConfig["profiles"][number]["profile"]["provider"],
): string => (isGeminiProvider(provider) ? "google" : "openai");

const getSupplierLabel = (
    provider: AIDetectConfig["profiles"][number]["profile"]["provider"],
): string =>
    provider === "modelrouter-openai" || provider === "modelrouter-gemini"
        ? "ModelRouter"
        : "Idealab";

const getProviderBySupplier = (
    supplier: string,
    provider: AIDetectConfig["profiles"][number]["profile"]["provider"],
): AIDetectConfig["profiles"][number]["profile"]["provider"] => {
    if (supplier === "ModelRouter") {
        return isGeminiProvider(provider)
            ? "modelrouter-gemini"
            : "modelrouter-openai";
    }
    return isGeminiProvider(provider) ? "gemini" : "openai";
};

const getProviderOptionsBySupplier = (
    supplier: string,
): readonly (typeof AI_PROVIDER_OPTIONS)[number][] =>
    AI_PROVIDER_OPTIONS.filter((option) =>
        supplier === "ModelRouter"
            ? option.value.startsWith("modelrouter-")
            : !option.value.startsWith("modelrouter-"),
    );

const splitModelId = (
    value: string,
    fallbackProvider: string,
): { provider: string; name: string } => {
    const trimmed = value.trim();
    if (!trimmed) {
        return { provider: fallbackProvider, name: "" };
    }
    const slashIndex = trimmed.indexOf("/");
    if (slashIndex > 0) {
        return {
            provider: trimmed.slice(0, slashIndex),
            name: trimmed.slice(slashIndex + 1),
        };
    }
    return { provider: fallbackProvider, name: trimmed };
};

const composeModelId = (_provider: string, name: string): string => {
    const trimmedName = name.trim();
    return trimmedName;
};

const stripModelProviderPrefix = (value: string): string => {
    const trimmed = value.trim();
    if (!trimmed) {
        return "";
    }
    const slashIndex = trimmed.indexOf("/");
    if (slashIndex <= 0) {
        return trimmed;
    }
    const prefix = trimmed.slice(0, slashIndex).toLowerCase();
    if (!MODEL_PROVIDER_PREFIXES.has(prefix)) {
        return trimmed;
    }
    const remainder = trimmed.slice(slashIndex + 1).trim();
    return remainder.length > 0 ? remainder : trimmed;
};

interface AIProfileModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    aiConfigFormMessage: string;
    draftAIConfig: AIDetectConfig;
    setDraftAIConfig: Dispatch<SetStateAction<AIDetectConfig>>;
    aiConfigSaving: boolean;
    onCancel: () => void;
    onSave: () => void;
}

const updateStageProfileName = (
    stages: AIDetectStageConfigMap,
    targetName: string,
    nextName: string,
): AIDetectStageConfigMap => {
    const updated: AIDetectStageConfigMap = { ...stages };
    AI_STAGE_ORDER.forEach((stageKey) => {
        const stage = stages[stageKey];
        updated[stageKey] =
            stage.profileName === targetName
                ? { ...stage, profileName: nextName }
                : stage;
    });
    return updated;
};

export function AIProfileModal({
    isOpen,
    activeFile,
    aiConfigFormMessage,
    draftAIConfig,
    setDraftAIConfig,
    aiConfigSaving,
    onCancel,
    onSave,
}: AIProfileModalProps) {
    if (!isOpen || !activeFile) {
        return null;
    }

    const profiles = draftAIConfig.profiles ?? [];

    const getNextProfileName = (base: string) => {
        if (!profiles.some((item) => item.name === base)) {
            return base;
        }
        let index = 2;
        while (profiles.some((item) => item.name === `${base} ${index}`)) {
            index += 1;
        }
        return `${base} ${index}`;
    };

    const onAddProfile = () => {
        const name = getNextProfileName("接口配置");
        const profile = cloneAIDetectProfile(DEFAULT_AI_PROFILE);
        setDraftAIConfig((previous) => ({
            ...previous,
            profiles: [...previous.profiles, { name, profile }],
        }));
    };

    const onRemoveProfile = (index: number) => {
        setDraftAIConfig((previous) => {
            if (previous.profiles.length <= 1) {
                return previous;
            }
            const removedName = previous.profiles[index]?.name ?? "";
            const nextProfiles = previous.profiles.filter(
                (_, itemIndex) => itemIndex !== index,
            );
            const fallbackName = nextProfiles[0]?.name ?? "";
            return {
                ...previous,
                profiles: nextProfiles,
                stages: removedName
                    ? updateStageProfileName(
                          previous.stages,
                          removedName,
                          fallbackName,
                      )
                    : previous.stages,
            };
        });
    };

    const onRenameProfile = (index: number, nextName: string) => {
        setDraftAIConfig((previous) => {
            const previousName = previous.profiles[index]?.name ?? "";
            const nextProfiles = previous.profiles.map((item, itemIndex) =>
                itemIndex === index ? { ...item, name: nextName } : item,
            );
            const nextStages =
                previousName && previousName !== nextName
                    ? updateStageProfileName(
                          previous.stages,
                          previousName,
                          nextName,
                      )
                    : previous.stages;
            return {
                ...previous,
                profiles: nextProfiles,
                stages: nextStages,
            };
        });
    };

    const updateProfile = (
        index: number,
        updater: (
            profile: AIDetectConfig["profiles"][number]["profile"],
        ) => AIDetectConfig["profiles"][number]["profile"],
    ) => {
        setDraftAIConfig((previous) => ({
            ...previous,
            profiles: previous.profiles.map((item, itemIndex) =>
                itemIndex === index
                    ? { ...item, profile: updater(item.profile) }
                    : item,
            ),
        }));
    };

    return (
        <div className="column-modal-mask">
            <div className="column-modal ai-config-modal">
                <h3>AI接口配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-profile-list">
                        {profiles.map((profileItem, index) => {
                            const profile = profileItem.profile;
                            const derivedModelParts = splitModelId(
                                profile.model,
                                getDefaultModelProvider(profile.provider),
                            );
                            const modelProvider =
                                profile.modelProvider?.trim() ||
                                derivedModelParts.provider;
                            const rawModelName =
                                profile.modelName?.trim() ||
                                derivedModelParts.name;
                            const modelName =
                                stripModelProviderPrefix(rawModelName);
                            const modelOptions =
                                MODEL_OPTIONS_BY_PROVIDER[modelProvider] ?? [];
                            const modelSelectOptions = modelOptions.includes(
                                modelName,
                            )
                                ? modelOptions
                                : modelName.trim().length > 0
                                  ? [modelName, ...modelOptions]
                                  : modelOptions;
                            const supplierLabel = getSupplierLabel(
                                profile.provider,
                            );
                            const providerOptions =
                                getProviderOptionsBySupplier(supplierLabel);
                            return (
                                <div key={index} className="ai-profile-card">
                                    <div className="ai-profile-card-head">
                                        <label className="ai-config-field ai-profile-card-title">
                                            <span>接口名称</span>
                                            <input
                                                type="text"
                                                value={profileItem.name}
                                                onChange={(event) =>
                                                    onRenameProfile(
                                                        index,
                                                        event.target.value,
                                                    )
                                                }
                                                placeholder="例如：默认接口 / 低成本接口"
                                            />
                                        </label>
                                        <div className="ai-profile-card-actions">
                                            <button
                                                type="button"
                                                className="btn btn-danger"
                                                onClick={() =>
                                                    onRemoveProfile(index)
                                                }
                                                disabled={profiles.length <= 1}
                                            >
                                                删除
                                            </button>
                                        </div>
                                    </div>
                                    <div className="ai-profile-card-grid">
                                        <label className="ai-config-field">
                                            <span>供应商</span>
                                            <select
                                                value={getSupplierLabel(
                                                    profile.provider,
                                                )}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => {
                                                            const nextProvider =
                                                                getProviderBySupplier(
                                                                    event.target
                                                                        .value,
                                                                    previous.provider,
                                                                );
                                                            const shouldSwitchToProviderDefaultUrl =
                                                                previous.url.trim()
                                                                    .length ===
                                                                    0 ||
                                                                previous.url.trim() ===
                                                                    getDefaultAIUrl(
                                                                        previous.provider,
                                                                    );
                                                            return {
                                                                ...previous,
                                                                provider:
                                                                    nextProvider,
                                                                url: shouldSwitchToProviderDefaultUrl
                                                                    ? getDefaultAIUrl(
                                                                          nextProvider,
                                                                      )
                                                                    : previous.url,
                                                                modelProvider:
                                                                    previous.modelProvider?.trim()
                                                                        .length
                                                                        ? previous.modelProvider
                                                                        : getDefaultModelProvider(
                                                                              nextProvider,
                                                                          ),
                                                            };
                                                        },
                                                    )
                                                }
                                            >
                                                <option value="Idealab">
                                                    Idealab
                                                </option>
                                                <option value="ModelRouter">
                                                    ModelRouter
                                                </option>
                                            </select>
                                        </label>
                                        <label className="ai-config-field">
                                            <span>接口类型</span>
                                            <select
                                                value={profile.provider}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => {
                                                            const nextProvider =
                                                                event.target
                                                                    .value as AIDetectConfig["profiles"][number]["profile"]["provider"];
                                                            const shouldSwitchToProviderDefaultUrl =
                                                                previous.url.trim()
                                                                    .length ===
                                                                    0 ||
                                                                previous.url.trim() ===
                                                                    getDefaultAIUrl(
                                                                        previous.provider,
                                                                    );
                                                            const fallbackModelProvider =
                                                                previous.modelProvider?.trim()
                                                                    .length
                                                                    ? previous.modelProvider
                                                                    : getDefaultModelProvider(
                                                                          nextProvider,
                                                                      );
                                                            const fallbackModelName =
                                                                previous.modelName?.trim()
                                                                    .length
                                                                    ? stripModelProviderPrefix(
                                                                          previous.modelName,
                                                                      )
                                                                    : (MODEL_OPTIONS_BY_PROVIDER[
                                                                          fallbackModelProvider
                                                                      ]?.[0] ??
                                                                      modelName);
                                                            return {
                                                                ...previous,
                                                                provider:
                                                                    nextProvider,
                                                                url: shouldSwitchToProviderDefaultUrl
                                                                    ? getDefaultAIUrl(
                                                                          nextProvider,
                                                                      )
                                                                    : previous.url,
                                                                modelProvider:
                                                                    fallbackModelProvider,
                                                                modelName:
                                                                    fallbackModelName,
                                                                model: fallbackModelName?.trim()
                                                                    .length
                                                                    ? composeModelId(
                                                                          fallbackModelProvider ??
                                                                              "",
                                                                          fallbackModelName ??
                                                                              "",
                                                                      )
                                                                    : previous.model,
                                                            };
                                                        },
                                                    )
                                                }
                                            >
                                                {providerOptions.map(
                                                    (option) => (
                                                        <option
                                                            key={option.value}
                                                            value={option.value}
                                                        >
                                                            {option.label}
                                                        </option>
                                                    ),
                                                )}
                                            </select>
                                        </label>
                                        <label className="ai-config-field">
                                            <span>Base URL</span>
                                            <input
                                                type="text"
                                                value={profile.url}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => ({
                                                            ...previous,
                                                            url: event.target
                                                                .value,
                                                        }),
                                                    )
                                                }
                                                placeholder={`例如：${getDefaultAIUrl(profile.provider)}`}
                                            />
                                        </label>
                                        <label className="ai-config-field">
                                            <span>模型提供商</span>
                                            <select
                                                value={modelProvider}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => {
                                                            const nextProvider =
                                                                event.target
                                                                    .value;
                                                            const previousModelName =
                                                                stripModelProviderPrefix(
                                                                    previous.modelName ??
                                                                        "",
                                                                );
                                                            const nextModelName =
                                                                MODEL_OPTIONS_BY_PROVIDER[
                                                                    nextProvider
                                                                ]?.includes(
                                                                    previousModelName,
                                                                )
                                                                    ? previousModelName
                                                                    : (MODEL_OPTIONS_BY_PROVIDER[
                                                                          nextProvider
                                                                      ]?.[0] ??
                                                                      previousModelName ??
                                                                      modelName);
                                                            return {
                                                                ...previous,
                                                                modelProvider:
                                                                    nextProvider,
                                                                modelName:
                                                                    nextModelName,
                                                                model: composeModelId(
                                                                    nextProvider,
                                                                    nextModelName ??
                                                                        "",
                                                                ),
                                                            };
                                                        },
                                                    )
                                                }
                                            >
                                                {MODEL_PROVIDER_OPTIONS.map(
                                                    (item) => (
                                                        <option
                                                            key={item}
                                                            value={item}
                                                        >
                                                            {item}
                                                        </option>
                                                    ),
                                                )}
                                            </select>
                                        </label>
                                        <label className="ai-config-field">
                                            <span>模型名称</span>
                                            <select
                                                value={modelName}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => {
                                                            const nextModelName =
                                                                stripModelProviderPrefix(
                                                                    event.target
                                                                        .value,
                                                                );
                                                            const nextModelProvider =
                                                                previous.modelProvider?.trim()
                                                                    .length
                                                                    ? previous.modelProvider
                                                                    : modelProvider;
                                                            return {
                                                                ...previous,
                                                                modelProvider:
                                                                    nextModelProvider,
                                                                modelName:
                                                                    nextModelName,
                                                                model: composeModelId(
                                                                    nextModelProvider ??
                                                                        "",
                                                                    nextModelName,
                                                                ),
                                                            };
                                                        },
                                                    )
                                                }
                                            >
                                                {modelSelectOptions.map(
                                                    (item) => (
                                                        <option
                                                            key={item}
                                                            value={item}
                                                        >
                                                            {item}
                                                        </option>
                                                    ),
                                                )}
                                            </select>
                                            {modelSelectOptions.length === 0 ? (
                                                <small className="ai-config-hint">
                                                    当前模型提供商没有可选模型
                                                </small>
                                            ) : null}
                                        </label>
                                        <label className="ai-config-field">
                                            <span>
                                                {isGeminiProvider(
                                                    profile.provider,
                                                )
                                                    ? "Thinking 级别"
                                                    : "Reasoning Effort"}
                                            </span>
                                            <select
                                                value={profile.reasoningEffort}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => ({
                                                            ...previous,
                                                            reasoningEffort:
                                                                event.target
                                                                    .value as AIDetectConfig["profiles"][number]["profile"]["reasoningEffort"],
                                                        }),
                                                    )
                                                }
                                            >
                                                {AI_REASONING_EFFORT_OPTIONS.map(
                                                    (option) => (
                                                        <option
                                                            key={option}
                                                            value={option}
                                                        >
                                                            {option}
                                                        </option>
                                                    ),
                                                )}
                                            </select>
                                        </label>
                                        <label className="ai-config-field">
                                            <span>失败重试次数（后端）</span>
                                            <input
                                                type="number"
                                                min={MIN_AI_RETRY_COUNT}
                                                max={MAX_AI_RETRY_COUNT}
                                                step={1}
                                                value={profile.retryCount}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => ({
                                                            ...previous,
                                                            retryCount:
                                                                normalizeAIRetryCount(
                                                                    Number(
                                                                        event
                                                                            .target
                                                                            .value,
                                                                    ),
                                                                ),
                                                        }),
                                                    )
                                                }
                                            />
                                        </label>
                                        <label className="ai-config-field">
                                            <span>接口 API Key</span>
                                            <input
                                                type="password"
                                                value={profile.apiKey}
                                                onChange={(event) =>
                                                    updateProfile(
                                                        index,
                                                        (previous) => ({
                                                            ...previous,
                                                            apiKey: event.target
                                                                .value,
                                                        }),
                                                    )
                                                }
                                                placeholder="请输入 API Key"
                                            />
                                        </label>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                    <div className="ai-profile-actions">
                        <button
                            type="button"
                            className="btn"
                            onClick={onAddProfile}
                        >
                            新增接口配置
                        </button>
                        <span className="ai-config-hint">
                            至少保留一个接口配置。
                        </span>
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
                        {aiConfigSaving ? "保存中..." : "保存AI接口配置"}
                    </button>
                </div>
            </div>
        </div>
    );
}
