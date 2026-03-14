import type { Express } from "express";
import fs from "node:fs";
import {
  getImageExtFromPathLike,
  getImageMimeType,
  LOCAL_IMAGE_API_PATH,
  toAbsoluteImagePath,
} from "../utils/images.js";

export const registerImageRoutes = (app: Express) => {
  app.get(LOCAL_IMAGE_API_PATH, (req, res) => {
    const pathQuery = req.query.path;
    if (typeof pathQuery !== "string" || pathQuery.trim().length === 0) {
      // eslint-disable-next-line no-console
      console.log("[ImageLocal] reject empty path query");
      return res.status(400).json({ message: "path is required" });
    }

    const absolutePath = toAbsoluteImagePath(pathQuery);
    if (!absolutePath) {
      // eslint-disable-next-line no-console
      console.log(`[ImageLocal] reject non-absolute path=${pathQuery}`);
      return res.status(400).json({ message: "path must be an absolute path" });
    }

    const ext = getImageExtFromPathLike(absolutePath);
    if (!ext) {
      // eslint-disable-next-line no-console
      console.log(
        `[ImageLocal] reject unsupported extension path=${absolutePath}`,
      );
      return res.status(400).json({ message: "unsupported image extension" });
    }

    try {
      if (!fs.existsSync(absolutePath) || !fs.statSync(absolutePath).isFile()) {
        // eslint-disable-next-line no-console
        console.log(`[ImageLocal] not found path=${absolutePath}`);
        return res.status(404).json({ message: "image not found" });
      }

      res.status(200);
      res.setHeader("Content-Type", getImageMimeType(ext));
      res.setHeader("Cache-Control", "public, max-age=120");
      const stream = fs.createReadStream(absolutePath);
      stream.on("error", () => {
        // eslint-disable-next-line no-console
        console.log(`[ImageLocal] read stream error path=${absolutePath}`);
        if (!res.headersSent) {
          res.status(500).json({ message: "read image failed" });
        } else {
          res.end();
        }
      });
      stream.pipe(res);
      return;
    } catch {
      // eslint-disable-next-line no-console
      console.log(`[ImageLocal] read image failed path=${absolutePath}`);
      return res.status(500).json({ message: "read image failed" });
    }
  });
};
