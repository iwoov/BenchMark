import {
    Fragment,
    useEffect,
    useRef,
    type KeyboardEvent,
    type ReactNode,
} from "react";
import type { AIModelRoute } from "../../types";
import type { AIChatMessage } from "../types";

interface AIChatSidebarProps {
    routes: AIModelRoute[];
    activeRouteName: string;
    chatMessages: AIChatMessage[];
    chatInput: string;
    chatStatusMessage: string;
    isAIChatting: boolean;
    aiChatElapsedText: string;
    onHide: () => void;
    onRouteChange: (value: string) => void;
    onInputChange: (value: string) => void;
    onSend: () => void;
    onClear: () => void;
}

function formatChatTime(timestamp: number): string {
    return new Date(timestamp).toLocaleTimeString("zh-CN", {
        hour: "2-digit",
        minute: "2-digit",
    });
}

function isSafeHttpUrl(value: string): boolean {
    try {
        const url = new URL(value);
        return url.protocol === "http:" || url.protocol === "https:";
    } catch {
        return false;
    }
}

function renderInlineMarkdown(text: string): ReactNode[] {
    const result: ReactNode[] = [];
    const pattern =
        /(\*\*[^*]+\*\*|`[^`]+`|\[[^\]]+\]\((https?:\/\/[^)\s]+)\))/g;
    let lastIndex = 0;
    let key = 0;

    for (const match of text.matchAll(pattern)) {
        const full = match[0];
        const index = match.index ?? 0;
        if (index > lastIndex) {
            result.push(text.slice(lastIndex, index));
        }

        if (full.startsWith("**") && full.endsWith("**")) {
            result.push(
                <strong key={`inline-${key}`}>
                    {full.slice(2, -2)}
                </strong>,
            );
            key += 1;
        } else if (full.startsWith("`") && full.endsWith("`")) {
            result.push(
                <code key={`inline-${key}`} className="detail-chat-inline-code">
                    {full.slice(1, -1)}
                </code>,
            );
            key += 1;
        } else if (full.startsWith("[")) {
            const linkMatch = /^\[([^\]]+)\]\((https?:\/\/[^)\s]+)\)$/.exec(full);
            if (linkMatch && isSafeHttpUrl(linkMatch[2])) {
                result.push(
                    <a
                        key={`inline-${key}`}
                        href={linkMatch[2]}
                        target="_blank"
                        rel="noreferrer"
                        className="detail-chat-link"
                    >
                        {linkMatch[1]}
                    </a>,
                );
                key += 1;
            } else {
                result.push(full);
            }
        }

        lastIndex = index + full.length;
    }

    if (lastIndex < text.length) {
        result.push(text.slice(lastIndex));
    }

    return result.length > 0 ? result : [text];
}

function renderMarkdownBlocks(content: string): ReactNode[] {
    const lines = content.replace(/\r\n?/g, "\n").split("\n");
    const blocks: ReactNode[] = [];
    let index = 0;

    while (index < lines.length) {
        const line = lines[index];
        const trimmed = line.trim();

        if (!trimmed) {
            index += 1;
            continue;
        }

        if (trimmed.startsWith("```")) {
            const codeLines: string[] = [];
            const fenceInfo = trimmed.slice(3).trim();
            index += 1;
            while (index < lines.length && !lines[index].trim().startsWith("```")) {
                codeLines.push(lines[index]);
                index += 1;
            }
            if (index < lines.length) {
                index += 1;
            }
            blocks.push(
                <pre className="detail-chat-code-block" key={`block-${blocks.length}`}>
                    {fenceInfo ? (
                        <span className="detail-chat-code-lang">{fenceInfo}</span>
                    ) : null}
                    <code>{codeLines.join("\n")}</code>
                </pre>,
            );
            continue;
        }

        const headingMatch = /^(#{1,6})\s+(.+)$/.exec(trimmed);
        if (headingMatch) {
            const level = Math.min(6, headingMatch[1].length);
            const HeadingTag = `h${level}` as keyof JSX.IntrinsicElements;
            blocks.push(
                <HeadingTag
                    className={`detail-chat-heading detail-chat-heading-${level}`}
                    key={`block-${blocks.length}`}
                >
                    {renderInlineMarkdown(headingMatch[2])}
                </HeadingTag>,
            );
            index += 1;
            continue;
        }

        if (trimmed.startsWith(">")) {
            const quoteLines: string[] = [];
            while (index < lines.length && lines[index].trim().startsWith(">")) {
                quoteLines.push(lines[index].trim().replace(/^>\s?/, ""));
                index += 1;
            }
            blocks.push(
                <blockquote
                    className="detail-chat-blockquote"
                    key={`block-${blocks.length}`}
                >
                    {quoteLines.map((quoteLine, quoteIndex) => (
                        <Fragment key={`quote-${quoteIndex}`}>
                            {quoteIndex > 0 ? <br /> : null}
                            {renderInlineMarkdown(quoteLine)}
                        </Fragment>
                    ))}
                </blockquote>,
            );
            continue;
        }

        if (/^[-*]\s+/.test(trimmed)) {
            const items: string[] = [];
            while (index < lines.length && /^[-*]\s+/.test(lines[index].trim())) {
                items.push(lines[index].trim().replace(/^[-*]\s+/, ""));
                index += 1;
            }
            blocks.push(
                <ul className="detail-chat-list" key={`block-${blocks.length}`}>
                    {items.map((item, itemIndex) => (
                        <li key={`item-${itemIndex}`}>
                            {renderInlineMarkdown(item)}
                        </li>
                    ))}
                </ul>,
            );
            continue;
        }

        if (/^\d+\.\s+/.test(trimmed)) {
            const items: string[] = [];
            while (index < lines.length && /^\d+\.\s+/.test(lines[index].trim())) {
                items.push(lines[index].trim().replace(/^\d+\.\s+/, ""));
                index += 1;
            }
            blocks.push(
                <ol className="detail-chat-list detail-chat-ordered-list" key={`block-${blocks.length}`}>
                    {items.map((item, itemIndex) => (
                        <li key={`item-${itemIndex}`}>
                            {renderInlineMarkdown(item)}
                        </li>
                    ))}
                </ol>,
            );
            continue;
        }

        const paragraphLines: string[] = [];
        while (index < lines.length) {
            const current = lines[index].trim();
            if (
                !current ||
                current.startsWith("```") ||
                current.startsWith(">") ||
                /^#{1,6}\s+/.test(current) ||
                /^[-*]\s+/.test(current) ||
                /^\d+\.\s+/.test(current)
            ) {
                break;
            }
            paragraphLines.push(lines[index]);
            index += 1;
        }
        blocks.push(
            <p className="detail-chat-paragraph" key={`block-${blocks.length}`}>
                {paragraphLines.map((paragraphLine, paragraphIndex) => (
                    <Fragment key={`paragraph-${paragraphIndex}`}>
                        {paragraphIndex > 0 ? <br /> : null}
                        {renderInlineMarkdown(paragraphLine)}
                    </Fragment>
                ))}
            </p>,
        );
    }

    return blocks;
}

function MarkdownMessageContent({ content }: { content: string }) {
    if (!content.trim()) {
        return <>AI 正在回答...</>;
    }
    return <>{renderMarkdownBlocks(content)}</>;
}

export function AIChatSidebar({
    routes,
    activeRouteName,
    chatMessages,
    chatInput,
    chatStatusMessage,
    isAIChatting,
    aiChatElapsedText,
    onHide,
    onRouteChange,
    onInputChange,
    onSend,
    onClear,
}: AIChatSidebarProps) {
    const messagesRef = useRef<HTMLDivElement | null>(null);

    useEffect(() => {
        const container = messagesRef.current;
        if (!container) {
            return;
        }
        container.scrollTop = container.scrollHeight;
    }, [chatMessages]);

    const onKeyDown = (event: KeyboardEvent<HTMLTextAreaElement>) => {
        if (event.key !== "Enter" || event.shiftKey) {
            return;
        }
        event.preventDefault();
        onSend();
    };

    return (
        <aside className="detail-chat-sidebar">
            <div className="detail-chat-header">
                <div className="detail-chat-title-group">
                    <strong>AI 聊天</strong>
                </div>
                <div className="detail-chat-header-actions">
                    <button
                        type="button"
                        className="btn btn-ghost detail-chat-new-btn"
                        onClick={onClear}
                        disabled={isAIChatting || chatMessages.length === 0}
                        aria-label="新建聊天"
                        title="新建聊天并清空当前会话记录"
                    >
                        +
                    </button>
                    <button
                        type="button"
                        className="btn btn-ghost detail-chat-toggle"
                        onClick={onHide}
                    >
                        收起
                    </button>
                </div>
            </div>

            <div className="detail-chat-messages" ref={messagesRef}>
                {chatMessages.length > 0 ? (
                    chatMessages.map((message) => (
                        <article
                            key={message.id}
                            className={`detail-chat-message role-${message.role} status-${message.status ?? "done"}`}
                        >
                            <div className="detail-chat-message-meta">
                                <strong>
                                    {message.role === "user" ? "你" : "AI"}
                                </strong>
                                <span>{formatChatTime(message.createdAt)}</span>
                            </div>
                            <div className="detail-chat-message-body">
                                <MarkdownMessageContent content={message.content} />
                            </div>
                        </article>
                    ))
                ) : (
                    <div className="detail-chat-empty">
                        这里会保留当前题目的本次会话记录。
                    </div>
                )}
            </div>

            <div className="detail-chat-composer">
                <textarea
                    value={chatInput}
                    onChange={(event) => onInputChange(event.target.value)}
                    onKeyDown={onKeyDown}
                    placeholder="输入你想追问的问题，Enter 发送，Shift+Enter 换行"
                    disabled={isAIChatting}
                />
                <div className="detail-chat-composer-actions">
                    <span className="detail-chat-status" role="status">
                        {chatStatusMessage ||
                            (isAIChatting
                                ? `AI 回答中 ${aiChatElapsedText}`
                                : "Enter 发送")}
                    </span>
                    <div className="detail-chat-action-group">
                        <label className="detail-chat-route detail-chat-route-inline">
                            <select
                                value={activeRouteName}
                                onChange={(event) =>
                                    onRouteChange(event.target.value)
                                }
                                disabled={isAIChatting || routes.length === 0}
                            >
                                {routes.length === 0 ? (
                                    <option value="">暂无可用路由</option>
                                ) : null}
                                {routes.map((route) => (
                                    <option key={route.name} value={route.name}>
                                        {route.name}
                                    </option>
                                ))}
                            </select>
                        </label>
                        <button
                            type="button"
                            className="btn btn-primary"
                            onClick={onSend}
                            disabled={
                                isAIChatting || chatInput.trim().length === 0
                            }
                        >
                            {isAIChatting ? "回答中..." : "发送"}
                        </button>
                    </div>
                </div>
            </div>
        </aside>
    );
}
