import type { Dispatch, SetStateAction } from "react";
import type {
  AIDetectConfig,
  AIDetectStageConfigMap,
  FileViewState,
  NamedAIDetectConfig,
} from "../../types";
import {
  AI_PROVIDER_OPTIONS,
  AI_REASONING_EFFORT_OPTIONS,
  AI_STAGE_ORDER,
  DEFAULT_AI_PROFILE,
  DEFAULT_GEMINI_URL,
  DEFAULT_OPENAI_URL,
  MAX_AI_RETRY_COUNT,
  MIN_AI_RETRY_COUNT,
} from "../constants";
import {
  cloneAIDetectProfile,
  getDefaultAIUrl,
  normalizeAIRetryCount,
} from "../ai-helpers";

interface AIProfileModalProps {
  isOpen: boolean;
  activeFile: FileViewState | null;
  aiConfigFormMessage: string;
  aiConfigList: NamedAIDetectConfig[];
  draftAIConfigName: string;
  setDraftAIConfigName: (value: string) => void;
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
  aiConfigList,
  draftAIConfigName,
  setDraftAIConfigName,
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
          ? updateStageProfileName(previous.stages, removedName, fallbackName)
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
          ? updateStageProfileName(previous.stages, previousName, nextName)
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
          <div className="column-modal-notice">{aiConfigFormMessage}</div>
        ) : null}
        <div className="ai-config-form">
          <label className="ai-config-field">
            <span>配置名称（输入新名称即新增）</span>
            <input
              type="text"
              value={draftAIConfigName}
              onChange={(event) => setDraftAIConfigName(event.target.value)}
              placeholder="例如：默认配置 / 低成本模型 / 高质量模型"
              list="ai-config-name-options"
            />
          </label>
          <datalist id="ai-config-name-options">
            {aiConfigList.map((item) => (
              <option key={item.name} value={item.name} />
            ))}
          </datalist>
          <div className="ai-profile-list">
            {profiles.map((profileItem, index) => {
              const profile = profileItem.profile;
              return (
                <div
                  key={`${profileItem.name}-${index}`}
                  className="ai-profile-card"
                >
                  <div className="ai-profile-card-head">
                    <label className="ai-config-field ai-profile-card-title">
                      <span>接口名称</span>
                      <input
                        type="text"
                        value={profileItem.name}
                        onChange={(event) =>
                          onRenameProfile(index, event.target.value)
                        }
                        placeholder="例如：默认接口 / 低成本接口"
                      />
                    </label>
                    <div className="ai-profile-card-actions">
                      <button
                        type="button"
                        className="btn btn-danger"
                        onClick={() => onRemoveProfile(index)}
                        disabled={profiles.length <= 1}
                      >
                        删除
                      </button>
                    </div>
                  </div>
                  <div className="ai-profile-card-grid">
                    <label className="ai-config-field">
                      <span>接口类型</span>
                      <select
                        value={profile.provider}
                        onChange={(event) =>
                          updateProfile(index, (previous) => {
                            const nextProvider = event.target
                              .value as AIDetectConfig["profiles"][number]["profile"]["provider"];
                            const shouldSwitchToProviderDefaultUrl =
                              previous.url.trim().length === 0 ||
                              previous.url.trim() ===
                                getDefaultAIUrl(previous.provider);
                            return {
                              ...previous,
                              provider: nextProvider,
                              url: shouldSwitchToProviderDefaultUrl
                                ? getDefaultAIUrl(nextProvider)
                                : previous.url,
                            };
                          })
                        }
                      >
                        {AI_PROVIDER_OPTIONS.map((option) => (
                          <option key={option.value} value={option.value}>
                            {option.label}
                          </option>
                        ))}
                      </select>
                    </label>
                    <label className="ai-config-field">
                      <span>
                        {profile.provider === "openai"
                          ? "OpenAI兼容接口 URL"
                          : "Gemini 接口 Endpoint"}
                      </span>
                      <input
                        type="text"
                        value={profile.url}
                        onChange={(event) =>
                          updateProfile(index, (previous) => ({
                            ...previous,
                            url: event.target.value,
                          }))
                        }
                        placeholder={
                          profile.provider === "openai"
                            ? `例如：${DEFAULT_OPENAI_URL}`
                            : `例如：${DEFAULT_GEMINI_URL}`
                        }
                      />
                    </label>
                    <label className="ai-config-field">
                      <span>模型</span>
                      <input
                        type="text"
                        value={profile.model}
                        onChange={(event) =>
                          updateProfile(index, (previous) => ({
                            ...previous,
                            model: event.target.value,
                          }))
                        }
                        placeholder={
                          profile.provider === "openai"
                            ? "例如：gpt-4.1-mini"
                            : "例如：gemini-2.5-flash"
                        }
                      />
                    </label>
                    <label className="ai-config-field">
                      <span>Reasoning 级别</span>
                      <select
                        value={profile.reasoningEffort}
                        onChange={(event) =>
                          updateProfile(index, (previous) => ({
                            ...previous,
                            reasoningEffort: event.target
                              .value as AIDetectConfig["profiles"][number]["profile"]["reasoningEffort"],
                          }))
                        }
                      >
                        {AI_REASONING_EFFORT_OPTIONS.map((option) => (
                          <option key={option} value={option}>
                            {option}
                          </option>
                        ))}
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
                          updateProfile(index, (previous) => ({
                            ...previous,
                            retryCount: normalizeAIRetryCount(
                              Number(event.target.value),
                            ),
                          }))
                        }
                      />
                    </label>
                    <label className="ai-config-field">
                      <span>
                        {profile.provider === "openai"
                          ? "OpenAI API Key"
                          : "Gemini API Key"}
                      </span>
                      <input
                        type="password"
                        value={profile.apiKey}
                        onChange={(event) =>
                          updateProfile(index, (previous) => ({
                            ...previous,
                            apiKey: event.target.value,
                          }))
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
            <button type="button" className="btn" onClick={onAddProfile}>
              新增接口配置
            </button>
            <span className="ai-config-hint">至少保留一个接口配置。</span>
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
