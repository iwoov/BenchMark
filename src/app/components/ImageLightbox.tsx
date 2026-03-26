import { useEffect, useState } from "react";

interface ImageLightboxProps {
    src: string | null;
    onClose: () => void;
}

export function ImageLightbox({ src, onClose }: ImageLightboxProps) {
    const [zoom, setZoom] = useState(100);

    useEffect(() => {
        setZoom(100);
    }, [src]);

    if (!src) {
        return null;
    }

    return (
        <div
            className="lightbox-mask"
            onClick={onClose}
            onKeyDown={(event) => {
                if (event.key === "Escape") {
                    onClose();
                }
            }}
            role="button"
            tabIndex={0}
        >
            <div
                className="lightbox-toolbar"
                onClick={(event) => event.stopPropagation()}
            >
                <button
                    type="button"
                    className="btn btn-ghost"
                    onClick={() =>
                        setZoom((previous) => Math.max(50, previous - 25))
                    }
                    disabled={zoom <= 50}
                >
                    -
                </button>
                <span>{`${zoom}%`}</span>
                <button
                    type="button"
                    className="btn btn-ghost"
                    onClick={() => setZoom(100)}
                >
                    100%
                </button>
                <button
                    type="button"
                    className="btn btn-ghost"
                    onClick={() =>
                        setZoom((previous) => Math.min(300, previous + 25))
                    }
                    disabled={zoom >= 300}
                >
                    +
                </button>
            </div>
            <img
                className="lightbox-image"
                src={src}
                alt="预览大图"
                onClick={(event) => event.stopPropagation()}
                style={{ transform: `scale(${zoom / 100})` }}
            />
            <button type="button" className="lightbox-close" onClick={onClose}>
                ×
            </button>
        </div>
    );
}
