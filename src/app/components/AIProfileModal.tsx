import type { Dispatch, SetStateAction } from "react";
import type { AIDetectConfig, FileViewState } from "../../types";
import {
    AI_PROVIDER_API_TYPE_OPTIONS,
    DEFAULT_AI_PROVIDER,
} from "../constants";
import {
    cloneAIProviderEndpoint,
    getDefaultProviderUrl,
} from "../ai-helpers";

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

    const providers = draftAIConfig.providers ?? [];

    const getNextProviderName = (base: string) => {
        if (!providers.some((item) => item.name === base)) {
            return base;
        }
        let index = 2;
        while (providers.some((item) => item.name === `${base} ${index}`)) {
            index += 1;
        }
        return `${base} ${index}`;
    };

    const onAddProvider = () => {
        const provider = cloneAIProviderEndpoint(DEFAULT_AI_PROVIDER);
        provider.name = getNextProviderName("模型提供商");
        setDraftAIConfig((previous) => ({
            ...previous,
            providers: [...previous.providers, provider],
        }));
    };

    const onRemoveProvider = (index: number) => {
        setDraftAIConfig((previous) => {
            if (previous.providers.length <= 1) {
                return previous;
            }
            const removedName = previous.providers[index]?.name ?? "";
            const nextProviders = previous.providers.filter(
                (_, itemIndex) => itemIndex !== index,
            );
            const fallbackProviderName = nextProviders[0]?.name ?? "";
            const nextRoutes = previous.routes.map((route) => ({
                ...route,
                steps:
                    route.steps.filter(
                        (step) => step.providerName !== removedName,
                    ).length > 0
                        ? route.steps.filter(
                              (step) => step.providerName !== removedName,
                          )
                        : fallbackProviderName
                          ? [{ providerName: fallbackProviderName }]
                          : [],
            }));
            return {
                ...previous,
                providers: nextProviders,
                routes: nextRoutes,
            };
        });
    };

    const updateProvider = (
        index: number,
        updater: (
            provider: AIDetectConfig["providers"][number],
        ) => AIDetectConfig["providers"][number],
    ) => {
        setDraftAIConfig((previous) => ({
            ...previous,
            providers: previous.providers.map((item, itemIndex) =>
                itemIndex === index ? updater(item) : item,
            ),
        }));
    };

    const onRenameProvider = (index: number, nextName: string) => {
        setDraftAIConfig((previous) => {
            const previousName = previous.providers[index]?.name ?? "";
            const normalizedName = nextName;
            return {
                ...previous,
                providers: previous.providers.map((item, itemIndex) =>
                    itemIndex === index ? { ...item, name: normalizedName } : item,
                ),
                routes:
                    previousName && previousName !== normalizedName
                        ? previous.routes.map((route) => ({
                              ...route,
                              steps: route.steps.map((step) =>
                                  step.providerName === previousName
                                      ? {
                                            ...step,
                                            providerName: normalizedName,
                                        }
                                      : step,
                              ),
                          }))
                        : previous.routes,
            };
        });
    };

    return (
        <div className="column-modal-mask">
            <div className="column-modal ai-config-modal">
                <h3>模型提供商配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-profile-list">
                        {providers.map((provider, index) => (
                            <div key={index} className="ai-profile-card">
                                <div className="ai-profile-card-head">
                                    <label className="ai-config-field ai-profile-card-title">
                                        <span>提供商名称</span>
                                        <input
                                            type="text"
                                            value={provider.name}
                                            onChange={(event) =>
                                                onRenameProvider(
                                                    index,
                                                    event.target.value,
                                                )
                                            }
                                            placeholder="例如：Idealab OpenAI / ModelRouter"
                                        />
                                    </label>
                                    <div className="ai-profile-card-actions">
                                        <button
                                            type="button"
                                            className="btn btn-danger"
                                            onClick={() => onRemoveProvider(index)}
                                            disabled={providers.length <= 1}
                                        >
                                            删除
                                        </button>
                                    </div>
                                </div>
                                <div className="ai-profile-card-grid">
                                    <label className="ai-config-field">
                                        <span>接口类型</span>
                                        <select
                                            value={provider.apiType}
                                            onChange={(event) =>
                                                updateProvider(index, (previous) => {
                                                    const apiType =
                                                        event.target
                                                            .value as AIDetectConfig["providers"][number]["apiType"];
                                                    const shouldUseDefaultUrl =
                                                        previous.apiUrl.trim()
                                                            .length === 0 ||
                                                        previous.apiUrl ===
                                                            getDefaultProviderUrl(
                                                                previous.apiType,
                                                            );
                                                    return {
                                                        ...previous,
                                                        apiType,
                                                        apiUrl: shouldUseDefaultUrl
                                                            ? getDefaultProviderUrl(
                                                                  apiType,
                                                              )
                                                            : previous.apiUrl,
                                                    };
                                                })
                                            }
                                        >
                                            {AI_PROVIDER_API_TYPE_OPTIONS.map(
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
                                        <span>API URL</span>
                                        <input
                                            type="text"
                                            value={provider.apiUrl}
                                            onChange={(event) =>
                                                updateProvider(index, (previous) => ({
                                                    ...previous,
                                                    apiUrl: event.target.value,
                                                }))
                                            }
                                            placeholder="请输入 provider API URL"
                                        />
                                    </label>
                                    <label className="ai-config-field">
                                        <span>API Key</span>
                                        <input
                                            type="password"
                                            value={provider.apiKey}
                                            onChange={(event) =>
                                                updateProvider(index, (previous) => ({
                                                    ...previous,
                                                    apiKey: event.target.value,
                                                }))
                                            }
                                            placeholder="请输入 API Key"
                                        />
                                    </label>
                                </div>
                            </div>
                        ))}
                    </div>
                    <div className="ai-profile-actions">
                        <button
                            type="button"
                            className="btn"
                            onClick={onAddProvider}
                        >
                            新增模型提供商
                        </button>
                        <span className="ai-config-hint">
                            provider 只维护接口类型、API URL 和 API Key。
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
                        {aiConfigSaving ? "保存中..." : "保存模型提供商"}
                    </button>
                </div>
            </div>
        </div>
    );
}
