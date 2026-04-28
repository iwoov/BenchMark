import { useEffect, useState } from "react";

interface ImageLightboxProps {
    src: string | null;
    srcList?: string[];
    onClose: () => void;
}

export function ImageLightbox({ src, srcList, onClose }: ImageLightboxProps) {
    const [zoom, setZoom] = useState(100);
    const images = srcList && srcList.length > 0 ? srcList : src ? [src] : [];
    const [activeIndex, setActiveIndex] = useState(0);

    useEffect(() => {
        setZoom(100);
        if (srcList && srcList.length > 0 && src) {
            const idx = srcList.indexOf(src);
            setActiveIndex(idx >= 0 ? idx : 0);
        } else {
            setActiveIndex(0);
        }
    }, [src, srcList]);

    if (images.length === 0) {
        return null;
    }

    const activeSrc = images[activeIndex] ?? images[0];
    const hasMultiple = images.length > 1;

    const goPrev = () => {
        setZoom(100);
        setActiveIndex((prev) => (prev - 1 + images.length) % images.length);
    };
    const goNext = () => {
        setZoom(100);
        setActiveIndex((prev) => (prev + 1) % images.length);
    };

    return (
        <div
            className="lightbox-mask"
            onClick={onClose}
            onKeyDown={(event) => {
                if (event.key === "Escape") {
                    onClose();
                }
                if (hasMultiple && event.key === "ArrowLeft") {
                    event.stopPropagation();
                    goPrev();
                }
                if (hasMultiple && event.key === "ArrowRight") {
                    event.stopPropagation();
                    goNext();
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
                {hasMultiple ? (
                    <span className="lightbox-counter">{`${activeIndex + 1} / ${images.length}`}</span>
                ) : null}
            </div>
            {hasMultiple ? (
                <button
                    type="button"
                    className="lightbox-nav lightbox-nav-prev"
                    onClick={(event) => {
                        event.stopPropagation();
                        goPrev();
                    }}
                    aria-label="上一张"
                >
                    ‹
                </button>
            ) : null}
            <img
                className="lightbox-image"
                src={activeSrc}
                alt="预览大图"
                onClick={(event) => event.stopPropagation()}
                style={{ transform: `scale(${zoom / 100})` }}
            />
            {hasMultiple ? (
                <button
                    type="button"
                    className="lightbox-nav lightbox-nav-next"
                    onClick={(event) => {
                        event.stopPropagation();
                        goNext();
                    }}
                    aria-label="下一张"
                >
                    ›
                </button>
            ) : null}
            <button type="button" className="lightbox-close" onClick={onClose}>
                ×
            </button>
        </div>
    );
}
