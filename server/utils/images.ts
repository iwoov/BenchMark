import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

export const LOCAL_IMAGE_API_PATH = "/api/images/local";
const SUPPORTED_IMAGE_EXTENSIONS = ["png", "jpg", "jpeg", "webp"] as const;

export function getImageMimeType(ext: string): string {
  const map: Record<string, string> = {
    png: "image/png",
    jpg: "image/jpeg",
    jpeg: "image/jpeg",
    webp: "image/webp",
  };
  return map[ext.toLowerCase()] || `image/${ext}`;
}

export function getImageExtFromPathLike(pathLike: string): string | null {
  const purePath = pathLike.split(/[?#]/)[0];
  const ext = path.extname(purePath).replace(".", "").toLowerCase();
  return SUPPORTED_IMAGE_EXTENSIONS.includes(
    ext as (typeof SUPPORTED_IMAGE_EXTENSIONS)[number],
  )
    ? ext
    : null;
}

export function toAbsoluteImagePath(pathLike: string): string | null {
  const trimmed = pathLike.trim();
  if (!trimmed) {
    return null;
  }

  if (/^file:\/\//i.test(trimmed)) {
    try {
      return fileURLToPath(new URL(trimmed));
    } catch {
      return null;
    }
  }

  if (path.isAbsolute(trimmed) || /^[a-zA-Z]:[\\/]/.test(trimmed)) {
    return trimmed;
  }

  return null;
}

export function toDataUrlFromAbsoluteImagePath(imagePath: string): string | null {
  const ext = getImageExtFromPathLike(imagePath);
  if (!ext) {
    return null;
  }

  try {
    if (!fs.existsSync(imagePath) || !fs.statSync(imagePath).isFile()) {
      return null;
    }
    const imageBuffer = fs.readFileSync(imagePath);
    return `data:${getImageMimeType(ext)};base64,${imageBuffer.toString("base64")}`;
  } catch {
    return null;
  }
}

export function tryGetPathFromLocalImageApiUrl(imageUrl: string): string | null {
  const trimmed = imageUrl.trim();
  if (!trimmed) {
    return null;
  }

  const parseByUrl = (urlLike: string): string | null => {
    try {
      const url = new URL(urlLike);
      if (url.pathname !== LOCAL_IMAGE_API_PATH) {
        return null;
      }
      const rawPath = url.searchParams.get("path");
      if (!rawPath) {
        return null;
      }
      return toAbsoluteImagePath(rawPath);
    } catch {
      return null;
    }
  };

  if (trimmed.startsWith(LOCAL_IMAGE_API_PATH)) {
    return parseByUrl(`http://localhost${trimmed}`);
  }

  if (/^https?:\/\//i.test(trimmed)) {
    return parseByUrl(trimmed);
  }

  return null;
}

export function normalizeImageUrlForAI(imageUrl: string): string | null {
  const trimmed = imageUrl.trim();
  if (!trimmed) {
    return null;
  }

  if (/^data:image\//i.test(trimmed)) {
    return trimmed;
  }

  const localPathFromApi = tryGetPathFromLocalImageApiUrl(trimmed);
  if (localPathFromApi) {
    return toDataUrlFromAbsoluteImagePath(localPathFromApi);
  }

  const absolutePath = toAbsoluteImagePath(trimmed);
  if (absolutePath) {
    return toDataUrlFromAbsoluteImagePath(absolutePath);
  }

  if (/^https?:\/\//i.test(trimmed)) {
    return trimmed;
  }

  return null;
}
