export type ToastKind = "success" | "error";

export interface ToastItem {
    id: string;
    kind: ToastKind;
    message: string;
}

interface ToastViewportProps {
    items: ToastItem[];
    onDismiss: (id: string) => void;
}

export function ToastViewport({
    items,
    onDismiss,
}: ToastViewportProps) {
    if (items.length === 0) {
        return null;
    }

    return (
        <div className="toast-viewport" aria-live="polite" aria-atomic="true">
            {items.map((item) => (
                <div
                    key={item.id}
                    className={`toast-card toast-${item.kind}`}
                    role="status"
                >
                    <div className="toast-card-indicator" />
                    <div className="toast-card-body">
                        <strong>
                            {item.kind === "success" ? "保存成功" : "保存失败"}
                        </strong>
                        <span>{item.message}</span>
                    </div>
                    <button
                        type="button"
                        className="toast-card-close"
                        onClick={() => onDismiss(item.id)}
                        aria-label="关闭通知"
                        title="关闭通知"
                    >
                        ×
                    </button>
                </div>
            ))}
        </div>
    );
}
