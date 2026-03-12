import type { FileViewState } from "../../types";
import type { MainSection, SettingsSection } from "../types";
import {
  IconInspect,
  IconListTree,
  IconPanelLeft,
  IconSliders,
} from "../icons";

interface WorkspaceSidebarProps {
  isCollapsed: boolean;
  activeSection: MainSection;
  activeSettingsSection: SettingsSection;
  activeFile: FileViewState | null;
  onToggle: () => void;
  onNavigate: (section: MainSection, settingsSection?: SettingsSection) => void;
}

export function WorkspaceSidebar({
  isCollapsed,
  activeSection,
  activeSettingsSection,
  activeFile,
  onToggle,
  onNavigate,
}: WorkspaceSidebarProps) {
  return (
    <aside className="workspace-sidebar">
      <div className="workspace-sidebar-header">
        <div className="workspace-sidebar-title">
          <span>工作区</span>
          {!isCollapsed ? <strong>页面导航</strong> : null}
        </div>
        <button
          type="button"
          className="workspace-sidebar-toggle"
          onClick={onToggle}
          aria-label={isCollapsed ? "展开左侧导航" : "折叠左侧导航"}
          title={isCollapsed ? "展开导航" : "折叠导航"}
        >
          <IconPanelLeft />
        </button>
      </div>
      <button
        type="button"
        className={`workspace-nav-item ${activeSection === "list" ? "active" : ""}`}
        onClick={() => onNavigate("list")}
        title="题目列表"
      >
        <span className="workspace-nav-icon" aria-hidden="true">
          <IconListTree />
        </span>
        <span className="workspace-nav-copy">
          <strong>题目列表</strong>
          <span>全字段浏览、筛选、分页与批量处理</span>
        </span>
      </button>
      <button
        type="button"
        className={`workspace-nav-item ${activeSection === "detail" ? "active" : ""}`}
        onClick={() => onNavigate("detail")}
        disabled={!activeFile}
        title="题目详情"
      >
        <span className="workspace-nav-icon" aria-hidden="true">
          <IconInspect />
        </span>
        <span className="workspace-nav-copy">
          <strong>题目详情</strong>
          <span>查看单条记录、字段编辑、AI 回答</span>
        </span>
      </button>
      <button
        type="button"
        className={`workspace-nav-item ${activeSection === "settings" ? "active" : ""}`}
        onClick={() => onNavigate("settings", activeSettingsSection)}
        disabled={!activeFile}
        title="设置区域"
      >
        <span className="workspace-nav-icon" aria-hidden="true">
          <IconSliders />
        </span>
        <span className="workspace-nav-copy">
          <strong>设置区域</strong>
          <span>字段规则与 AI 配置入口</span>
        </span>
      </button>
    </aside>
  );
}
