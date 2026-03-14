import cors from "cors";
import express from "express";
import multer from "multer";
import { registerAIRoutes } from "./ai/index.js";
import { registerExportRoute } from "./routes/export.js";
import { registerFileRoutes } from "./routes/files.js";
import { registerHealthRoute } from "./routes/health.js";
import { registerImageRoutes } from "./routes/images.js";

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

app.listen(port, () => {
  // eslint-disable-next-line no-console
  console.log(`Server running at http://localhost:${port}`);
});
