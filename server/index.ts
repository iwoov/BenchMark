import cors from "cors";
import express from "express";
import multer from "multer";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { registerAIRoutes } from "./ai/index.js";
import { registerExportRoute } from "./routes/export.js";
import { registerFileRoutes } from "./routes/files.js";
import { registerHealthRoute } from "./routes/health.js";
import { registerImageRoutes } from "./routes/images.js";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const clientDistDir = path.resolve(__dirname, "..", "dist");
const clientIndexHtmlPath = path.join(clientDistDir, "index.html");
const app = express();
const port = 8787;

const upload = multer({
    storage: multer.memoryStorage(),
    limits: {
        fileSize: 25 * 1024 * 1024,
    },
});

app.use(cors());
app.use(express.json({ limit: "100mb" }));

registerHealthRoute(app);
registerImageRoutes(app);
registerFileRoutes(app, upload);
registerExportRoute(app);
registerAIRoutes(app);

if (fs.existsSync(clientDistDir) && fs.existsSync(clientIndexHtmlPath)) {
    app.use(express.static(clientDistDir));
    app.get("*", (_req, res) => {
        res.sendFile(clientIndexHtmlPath);
    });
}

app.listen(port, () => {
    // eslint-disable-next-line no-console
    console.log(`Server running at http://localhost:${port}`);
});
