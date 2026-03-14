import type { Express } from "express";

export const registerHealthRoute = (app: Express) => {
  app.get("/api/health", (_req, res) => {
    res.json({
      ok: true,
      timestamp: new Date().toISOString(),
    });
  });
};
