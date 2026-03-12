interface ImageLightboxProps {
  src: string | null;
  onClose: () => void;
}

export function ImageLightbox({ src, onClose }: ImageLightboxProps) {
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
      <img
        className="lightbox-image"
        src={src}
        alt="预览大图"
        onClick={(event) => event.stopPropagation()}
      />
      <button type="button" className="lightbox-close" onClick={onClose}>
        ×
      </button>
    </div>
  );
}
