import type {
  AIDetectConfig,
  AIDetectProfile,
  FileViewState,
  NamedAIDetectConfig,
  ParsedColumn,
} from "../../types";
import {
  AI_PROVIDER_OPTIONS,
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
}: SettingsPageProps) {
  const visibleDisplayColumns = displayColumns;
  const editableColumns = activeFile.columns.filter((column) =>
    activeFile.selectedEditableColumnKeys.includes(column.key),
  );
  const filterColumns = activeFile.columns.filter((column) =>
    activeFile.selectedFilterColumnKeys.includes(column.key),
  );
  const activeConfig = aiConfigList[0]?.config ?? aiConfig;
  const profiles = activeConfig.profiles ?? [];
  const profileMap = new Map(profiles.map((item) => [item.name, item.profile]));
  const openaiConfigs = profiles.filter(
    (item) => item.profile.provider === "openai",
  );
  const geminiConfigs = profiles.filter(
    (item) => item.profile.provider === "gemini",
  );
  const getProviderLabel = (provider: AIDetectProfile["provider"]) =>
    AI_PROVIDER_OPTIONS.find((item) => item.value === provider)?.label ?? "-";
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
              <p>统一控制详情页哪些字段显示，以及哪些字段允许编辑。</p>
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
                {visibleDisplayColumns.slice(0, 12).map((column) => (
                  <span key={column.key} className="settings-pill">
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
                  editableColumns.slice(0, 12).map((column) => (
                    <span key={column.key} className="settings-pill">
                      {column.title}
                    </span>
                  ))
                ) : (
                  <span className="settings-empty">暂无可编辑字段</span>
                )}
                {editableColumns.length > 12 ? (
                  <span className="settings-pill">{`还有 ${
                    editableColumns.length - 12
                  } 个字段`}</span>
                ) : null}
              </div>
            </div>
            <div className="settings-subsection">
              <div className="settings-subsection-head">
                <h4>筛选字段</h4>
                <span>{`已选择 ${filterColumns.length} 个`}</span>
              </div>
              <div className="settings-pill-list">
                {filterColumns.length > 0 ? (
                  filterColumns.slice(0, 12).map((column) => (
                    <span key={column.key} className="settings-pill">
                      {column.title}
                    </span>
                  ))
                ) : (
                  <span className="settings-empty">暂无筛选字段</span>
                )}
                {filterColumns.length > 12 ? (
                  <span className="settings-pill">{`还有 ${
                    filterColumns.length - 12
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
            <p>配置不同阶段任务的提交字段/提示词，并维护模型接口。</p>
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
                  const profile = profileMap.get(stageConfig.profileName);
                  const providerLabel = profile
                    ? getProviderLabel(profile.provider)
                    : "未配置接口";
                  return (
                    <div key={stageKey} className="settings-stage-item">
                      <div className="settings-stage-title">
                        <strong>{stageLabel.title}</strong>
                        <span className="settings-tag">
                          {stageConfig.profileName || "未绑定接口"}
                        </span>
                        <span className="settings-tag">{providerLabel}</span>
                      </div>
                      <div className="settings-stage-meta">
                        <span>{`提交字段 ${stageConfig.submitFieldKeys.length} 个`}</span>
                        <span>{`结果字段 AI检测结果`}</span>
                        <span>{`重试 ${profile?.retryCount ?? 0} 次`}</span>
                        <span>{`模型 ${profile?.model || "-"}`}</span>
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
                <h4>模型接口配置</h4>
                <span>{`共 ${profiles.length} 个接口`}</span>
              </div>
              <div className="settings-config-grid">
                <div className="settings-config-group">
                  <h5>OpenAI 兼容 (Idealab)</h5>
                  {openaiConfigs.length > 0 ? (
                    openaiConfigs.map((item) => {
                      return (
                        <div key={item.name} className="settings-config-item">
                          <strong>{item.name}</strong>
                          <span>{item.profile.model || "未配置模型"}</span>
                          <span className="settings-config-url">
                            {item.profile.url || "未配置接口 URL"}
                          </span>
                        </div>
                      );
                    })
                  ) : (
                    <span className="settings-empty">暂无配置</span>
                  )}
                </div>
                <div className="settings-config-group">
                  <h5>Gemini 原生 (Idealab)</h5>
                  {geminiConfigs.length > 0 ? (
                    geminiConfigs.map((item) => {
                      return (
                        <div key={item.name} className="settings-config-item">
                          <strong>{item.name}</strong>
                          <span>{item.profile.model || "未配置模型"}</span>
                          <span className="settings-config-url">
                            {item.profile.url || "未配置接口 URL"}
                          </span>
                        </div>
                      );
                    })
                  ) : (
                    <span className="settings-empty">暂无配置</span>
                  )}
                </div>
              </div>
              <div className="settings-section-actions">
                <button
                  type="button"
                  className="btn"
                  onClick={onOpenAIProfileModal}
                >
                  管理接口配置
                </button>
              </div>
            </div>
          </div>
        </section>
      ) : null}
    </div>
  );
}
