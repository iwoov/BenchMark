import { useState } from "react";
import type { Dispatch, SetStateAction } from "react";
import type { AIDetectConfig, FileViewState } from "../../types";
import {
    AI_REASONING_EFFORT_OPTIONS,
    DEFAULT_AI_ROUTE,
} from "../constants";
import { cloneAIModelRoute, isAnthropicApiType } from "../ai-helpers";

interface AIRouteModalProps {
    isOpen: boolean;
    activeFile: FileViewState | null;
    aiConfigFormMessage: string;
    draftAIConfig: AIDetectConfig;
    setDraftAIConfig: Dispatch<SetStateAction<AIDetectConfig>>;
    aiConfigSaving: boolean;
    onCancel: () => void;
    onSave: () => void;
}

export function AIRouteModal({
    isOpen,
    activeFile,
    aiConfigFormMessage,
    draftAIConfig,
    setDraftAIConfig,
    aiConfigSaving,
    onCancel,
    onSave,
}: AIRouteModalProps) {
    const [testingKey, setTestingKey] = useState("");
    const [testMessage, setTestMessage] = useState("");

    if (!isOpen || !activeFile) {
        return null;
    }

    const routes = draftAIConfig.routes ?? [];
    const providers = draftAIConfig.providers ?? [];
    const providerMap = new Map(providers.map((item) => [item.name, item]));

    const getNextRouteName = (base: string) => {
        if (!routes.some((item) => item.name === base)) {
            return base;
        }
        let index = 2;
        while (routes.some((item) => item.name === `${base} ${index}`)) {
            index += 1;
        }
        return `${base} ${index}`;
    };

    const updateRoute = (
        index: number,
        updater: (
            route: AIDetectConfig["routes"][number],
        ) => AIDetectConfig["routes"][number],
    ) => {
        setDraftAIConfig((previous) => ({
            ...previous,
            routes: previous.routes.map((item, itemIndex) =>
                itemIndex === index ? updater(item) : item,
            ),
        }));
    };

    const onAddRoute = () => {
        const route = cloneAIModelRoute(DEFAULT_AI_ROUTE);
        route.name = getNextRouteName("模型路由");
        route.steps =
            providers[0]?.name ? [{ providerName: providers[0].name }] : [];
        setDraftAIConfig((previous) => ({
            ...previous,
            routes: [...previous.routes, route],
        }));
    };

    const onRemoveRoute = (index: number) => {
        setDraftAIConfig((previous) => ({
            ...previous,
            routes:
                previous.routes.length <= 1
                    ? previous.routes
                    : previous.routes.filter((_, itemIndex) => itemIndex !== index),
            stages:
                previous.routes.length <= 1
                    ? previous.stages
                    : Object.fromEntries(
                          Object.entries(previous.stages).map(([stageKey, stage]) => [
                              stageKey,
                              stage.routeName === previous.routes[index]?.name
                                  ? {
                                        ...stage,
                                        routeName:
                                            previous.routes.find(
                                                (_, itemIndex) => itemIndex !== index,
                                            )?.name ?? "",
                                    }
                                  : stage,
                          ]),
                      ) as AIDetectConfig["stages"],
        }));
    };

    const onRenameRoute = (index: number, nextName: string) => {
        setDraftAIConfig((previous) => {
            const previousName = previous.routes[index]?.name ?? "";
            return {
                ...previous,
                routes: previous.routes.map((item, itemIndex) =>
                    itemIndex === index ? { ...item, name: nextName } : item,
                ),
                stages:
                    previousName && previousName !== nextName
                        ? Object.fromEntries(
                              Object.entries(previous.stages).map(
                                  ([stageKey, stage]) => [
                                      stageKey,
                                      stage.routeName === previousName
                                          ? { ...stage, routeName: nextName }
                                          : stage,
                                  ],
                              ),
                          ) as AIDetectConfig["stages"]
                        : previous.stages,
            };
        });
    };

    const onAddStep = (index: number) => {
        if (!providers[0]?.name) {
            return;
        }
        updateRoute(index, (previous) => ({
            ...previous,
            steps: [...previous.steps, { providerName: providers[0].name }],
        }));
    };

    const onRemoveStep = (routeIndex: number, stepIndex: number) => {
        updateRoute(routeIndex, (previous) => ({
            ...previous,
            steps:
                previous.steps.length <= 1
                    ? previous.steps
                    : previous.steps.filter((_, index) => index !== stepIndex),
        }));
    };

    const onTestStep = async (routeIndex: number, stepIndex: number) => {
        const route = routes[routeIndex];
        const step = route?.steps[stepIndex];
        const provider = step ? providerMap.get(step.providerName) : null;
        if (!route || !step || !provider) {
            setTestMessage("当前步骤引用的模型提供商不存在");
            return;
        }
        if (isAnthropicApiType(provider.apiType)) {
            setTestMessage("Anthropic 暂不支持测试");
            return;
        }
        const currentKey = `${routeIndex}-${stepIndex}`;
        setTestingKey(currentKey);
        setTestMessage("");
        try {
            const response = await fetch("/api/ai-config/routes/test", {
                method: "POST",
                headers: {
                    "Content-Type": "application/json",
                },
                body: JSON.stringify({
                    provider,
                    route: {
                        ...route,
                        steps: [{ providerName: provider.name }],
                    },
                    stepIndex: 0,
                }),
            });
            const payload = (await response.json().catch(() => ({}))) as {
                message?: string;
                preview?: string;
                durationMs?: number;
            };
            if (!response.ok) {
                throw new Error(payload.message ?? "测试失败");
            }
            setTestMessage(
                `测试成功（${payload.durationMs ?? 0}ms）${
                    payload.preview ? `：${payload.preview}` : ""
                }`,
            );
        } catch (error) {
            setTestMessage(error instanceof Error ? error.message : "测试失败");
        } finally {
            setTestingKey("");
        }
    };

    return (
        <div className="column-modal-mask">
            <div className="column-modal ai-config-modal">
                <h3>模型路由配置</h3>
                <p>{activeFile.fileName}</p>
                {aiConfigFormMessage ? (
                    <div className="column-modal-notice">
                        {aiConfigFormMessage}
                    </div>
                ) : null}
                {testMessage ? (
                    <div className="column-modal-notice">{testMessage}</div>
                ) : null}
                <div className="ai-config-form">
                    <div className="ai-profile-list">
                        {routes.map((route, routeIndex) => (
                            <div key={routeIndex} className="ai-profile-card">
                                <div className="ai-profile-card-head">
                                    <label className="ai-config-field ai-profile-card-title">
                                        <span>路由名称</span>
                                        <input
                                            type="text"
                                            value={route.name}
                                            onChange={(event) =>
                                                onRenameRoute(
                                                    routeIndex,
                                                    event.target.value,
                                                )
                                            }
                                            placeholder="例如：gpt-5.4 / 低成本路由"
                                        />
                                    </label>
                                    <div className="ai-profile-card-actions">
                                        <button
                                            type="button"
                                            className="btn btn-danger"
                                            onClick={() => onRemoveRoute(routeIndex)}
                                            disabled={routes.length <= 1}
                                        >
                                            删除
                                        </button>
                                    </div>
                                </div>
                                <div className="ai-profile-card-grid">
                                    <label className="ai-config-field">
                                        <span>模型名称</span>
                                        <input
                                            type="text"
                                            value={route.model}
                                            onChange={(event) =>
                                                updateRoute(routeIndex, (previous) => ({
                                                    ...previous,
                                                    model: event.target.value,
                                                }))
                                            }
                                            placeholder="例如：gpt-5.4-2026-03-05"
                                        />
                                    </label>
                                    <label className="ai-config-field">
                                        <span>Reasoning / Thinking</span>
                                        <select
                                            value={route.reasoningEffort}
                                            onChange={(event) =>
                                                updateRoute(routeIndex, (previous) => ({
                                                    ...previous,
                                                    reasoningEffort:
                                                        event.target
                                                            .value as AIDetectConfig["routes"][number]["reasoningEffort"],
                                                }))
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
                                        <span>失败重试次数</span>
                                        <input
                                            type="number"
                                            min={0}
                                            max={10}
                                            step={1}
                                            value={route.retryCount}
                                            onChange={(event) =>
                                                updateRoute(routeIndex, (previous) => ({
                                                    ...previous,
                                                    retryCount: Number(
                                                        event.target.value,
                                                    ),
                                                }))
                                            }
                                        />
                                    </label>
                                </div>

                                <div className="ai-config-section">
                                    <div className="ai-config-section-title">
                                        回退步骤
                                    </div>
                                    <div className="ai-config-fields">
                                        {route.steps.map((step, stepIndex) => {
                                            const provider = providerMap.get(
                                                step.providerName,
                                            );
                                            const isTesting =
                                                testingKey ===
                                                `${routeIndex}-${stepIndex}`;
                                            return (
                                                <div
                                                    key={`${routeIndex}-${stepIndex}`}
                                                    className="ai-config-field-item"
                                                >
                                                    <select
                                                        value={step.providerName}
                                                        onChange={(event) => {
                                                            updateRoute(
                                                                routeIndex,
                                                                (previous) => ({
                                                                    ...previous,
                                                                    steps: previous.steps.map(
                                                                        (
                                                                            item,
                                                                            itemIndex,
                                                                        ) =>
                                                                            itemIndex ===
                                                                            stepIndex
                                                                                ? {
                                                                                      ...item,
                                                                                      providerName:
                                                                                          event.target.value,
                                                                                  }
                                                                                : item,
                                                                    ),
                                                                }),
                                                            );
                                                        }}
                                                    >
                                                        {providers.map((item) => (
                                                            <option
                                                                key={item.name}
                                                                value={item.name}
                                                            >
                                                                {item.name}
                                                            </option>
                                                        ))}
                                                    </select>
                                                    <button
                                                        type="button"
                                                        className="btn"
                                                        onClick={() =>
                                                            onTestStep(
                                                                routeIndex,
                                                                stepIndex,
                                                            )
                                                        }
                                                        disabled={
                                                            !provider ||
                                                            isAnthropicApiType(
                                                                provider.apiType,
                                                            ) ||
                                                            isTesting
                                                        }
                                                    >
                                                        {isTesting
                                                            ? "测试中..."
                                                            : "测试"}
                                                    </button>
                                                    <button
                                                        type="button"
                                                        className="btn btn-danger"
                                                        onClick={() =>
                                                            onRemoveStep(
                                                                routeIndex,
                                                                stepIndex,
                                                            )
                                                        }
                                                        disabled={
                                                            route.steps.length <= 1
                                                        }
                                                    >
                                                        删除步骤
                                                    </button>
                                                </div>
                                            );
                                        })}
                                    </div>
                                    <div className="ai-profile-actions">
                                        <button
                                            type="button"
                                            className="btn"
                                            onClick={() => onAddStep(routeIndex)}
                                            disabled={providers.length === 0}
                                        >
                                            新增回退步骤
                                        </button>
                                        <span className="ai-config-hint">
                                            首包返回前失败会自动切换到下一级提供商。
                                        </span>
                                    </div>
                                </div>
                            </div>
                        ))}
                    </div>
                    <div className="ai-profile-actions">
                        <button type="button" className="btn" onClick={onAddRoute}>
                            新增模型路由
                        </button>
                        <span className="ai-config-hint">
                            当前路由参数会应用到整条回退链路。
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
                        {aiConfigSaving ? "保存中..." : "保存模型路由"}
                    </button>
                </div>
            </div>
        </div>
    );
}
