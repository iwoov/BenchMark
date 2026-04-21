# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
pnpm dev              # Start frontend (port 5174) + backend (port 8787) concurrently
pnpm dev:web          # Frontend only (Vite)
pnpm dev:server       # Backend only (tsx watch)
pnpm typecheck        # TypeScript type checking (both client and server tsconfig)
pnpm build            # typecheck + build:server + vite build
pnpm test:server      # Run server unit tests (node --import tsx --test server/jsonParser.test.ts)
```

## Architecture

Full-stack TypeScript app: React SPA frontend + Express backend + SQLite.

### Frontend (`src/`)

- **Single entry**: `src/App.tsx` — monolithic component that owns all modal/sidebar state, wired to sub-hooks
- **Routing**: Hash-based (`#/dashboard`, `#/list`, `#/list/:rowId`, `#/settings/ai`, `#/projects`), parsed/built in `src/app/routes.ts`
- **State**: `useFileStore` (`src/app/hooks/useFileStore.ts`) is the central store — manages all loaded files, active file, upload/export, AI run state, and column config
- **Pages**: `DashboardPage`, `ListPage`, `DetailPage`, `SettingsPage`, `ProjectManagementPage` in `src/app/components/`
- **AI state**: `useAIManager` handles streaming AI evaluation runs and chat sidebar state

### Backend (`server/`)

- **Entry**: `server/index.ts` — Express app registering route modules
- **Database**: `server/db.ts` — single SQLite file at `data/benchmark.db` (WAL mode). Stores: `column_prefs`, `file_states`, AI provider endpoints, AI model routes, per-file AI stage/cleaning/chat/evaluation configs
- **Routes**: `server/routes/` — files (upload/list/fetch), export (Excel), images (filesystem image serving), health
- **AI**: `server/ai/index.ts` — multi-stage evaluation pipeline and streaming chat. Supports OpenAI, Gemini, Anthropic APIs via configurable provider endpoints and named routes

### AI Evaluation Pipeline

4 sequential stages per row: `precheck` → `context_audit` → `independent_solving` → `final_verdict`. Each stage is independently configurable (prompt, model route, enabled/disabled). Results are stored as JSON in `file_states`.

### Data Model

Files are identified by a stable `fileId` (UUID). File states (rows, column config, AI results) are persisted in SQLite as `state_json`. Images are served from filesystem paths configured via the `BENCHMARK_IMAGE_ROOTS` env var.

### Import Formats

- **Excel**: `.xls`/`.xlsx` parsed server-side via `exceljs`
- **JSON**: Two formats — workbench spec `{ columns, rows }` or plain object array `[{...}]`
- **Projects**: Multiple source files can be merged into one project file

### Key env vars

- `PORT` — server port (default `8787`)
- `HOST` — bind host (default `0.0.0.0`)
- `BENCHMARK_IMAGE_ROOTS` — colon-separated filesystem paths for resolving relative image paths
