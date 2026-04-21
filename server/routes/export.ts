import type { Express } from "express";
import * as ExcelJS from "exceljs";
import {
    getFileState,
    getFileAIEvaluationConfig,
    listFileAIEvaluationResults,
    listAIModelRoutes,
    listAIProviderEndpoints,
} from "../db.js";

type ExcelJSImportLike = { Workbook: new () => ExcelJS.Workbook };
const ExcelJSRuntime: ExcelJSImportLike =
    (ExcelJS as unknown as { default?: ExcelJSImportLike }).default ??
    (ExcelJS as unknown as ExcelJSImportLike);

export const registerExportRoute = (app: Express) => {
    app.post("/api/files/export", async (req, res) => {
        const { fileName, headers, rows } = req.body as {
            fileName: unknown;
            headers: unknown;
            rows: unknown;
        };

        if (typeof fileName !== "string" || fileName.trim().length === 0) {
            return res
                .status(400)
                .json({ message: "fileName must be a non-empty string" });
        }
        if (
            !Array.isArray(headers) ||
            !headers.every((item) => typeof item === "string")
        ) {
            return res
                .status(400)
                .json({ message: "headers must be a string array" });
        }
        if (
            !Array.isArray(rows) ||
            !rows.every(
                (row) =>
                    Array.isArray(row) &&
                    row.every((cell) => typeof cell === "string"),
            )
        ) {
            return res
                .status(400)
                .json({ message: "rows must be a 2d string array" });
        }

        try {
            const workbook = new ExcelJSRuntime.Workbook();
            const worksheet = workbook.addWorksheet("Sheet1");

            worksheet.addRow(headers);
            for (const row of rows) {
                worksheet.addRow(row);
            }

            worksheet.columns = headers.map((header, index) => {
                const maxLengthFromRows = rows.reduce((acc, row) => {
                    const value = row[index] ?? "";
                    return Math.max(acc, value.length);
                }, 0);
                return {
                    header,
                    key: `col_${index}`,
                    width: Math.min(
                        60,
                        Math.max(
                            12,
                            Math.max(header.length, maxLengthFromRows) + 2,
                        ),
                    ),
                };
            });

            const headerRow = worksheet.getRow(1);
            headerRow.font = { bold: true };
            headerRow.commit();

            const baseName = fileName.replace(/\.[^.]+$/, "");
            const exportName = `${baseName}-导出.xlsx`;
            const encodedFileName = encodeURIComponent(exportName);
            const buffer = await workbook.xlsx.writeBuffer();

            res.setHeader(
                "Content-Type",
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            );
            res.setHeader(
                "Content-Disposition",
                `attachment; filename*=UTF-8''${encodedFileName}`,
            );
            const outputBuffer = Buffer.isBuffer(buffer)
                ? buffer
                : Buffer.from(buffer as ArrayBuffer);
            return res.send(outputBuffer);
        } catch (error) {
            const message =
                error instanceof Error ? error.message : "导出 Excel 失败";
            return res.status(500).json({ message });
        }
    });

    app.post("/api/files/:fileId/export-evaluation", (req, res) => {
        const { fileId } = req.params;
        const { rowIds } = req.body as { rowIds?: unknown };

        const persisted = getFileState(fileId);
        if (!persisted) {
            return res.status(404).json({ message: "文件不存在" });
        }

        const state = persisted.state as {
            fileId?: string;
            fileName?: string;
            sourceFileName?: string;
            projectId?: string;
            dataSourceName?: string;
            columns?: Array<{ key: string; title: string }>;
            rows?: Array<{
                rowId: string;
                enabled?: boolean;
                values?: Record<string, { value?: string }>;
            }>;
        } | null;
        if (!state || !Array.isArray(state.rows)) {
            return res.status(404).json({ message: "文件数据为空" });
        }

        // Use client-provided rowIds if non-empty; otherwise export all rows
        const clientRowIds =
            Array.isArray(rowIds) &&
            rowIds.length > 0 &&
            rowIds.every((id) => typeof id === "string")
                ? (rowIds as string[])
                : null;
        const enabledRowIds: Set<string> | null = clientRowIds
            ? new Set(clientRowIds)
            : null;

        const enabledRows = enabledRowIds
            ? state.rows.filter((r) => enabledRowIds.has(r.rowId))
            : state.rows;

        const fileName = state.fileName ?? persisted.fileName;
        const evaluationResultsByTask = listFileAIEvaluationResults(fileId);
        if (Object.keys(evaluationResultsByTask).length === 0) {
            return res.status(404).json({ message: "暂无评测结果" });
        }

        const evaluationTasks = getFileAIEvaluationConfig(fileName);
        const routes = listAIModelRoutes();
        const providers = listAIProviderEndpoints();
        const routesByName = new Map(routes.map((r) => [r.name, r]));
        const providersByName = new Map(providers.map((p) => [p.name, p]));
        const columns = state.columns ?? [];
        const columnTitleByKey = new Map(columns.map((c) => [c.key, c.title]));
        const rowById = new Map(enabledRows.map((r) => [r.rowId, r]));
        const taskConfigById = new Map(evaluationTasks.map((t) => [t.id, t]));

        const exportItems: Record<string, unknown>[] = [];

        for (const [taskId, rowMap] of Object.entries(
            evaluationResultsByTask,
        )) {
            const taskConfig = taskConfigById.get(taskId);

            const genRouteName = taskConfig?.answerGeneration.routeName;
            const judgeRouteName = taskConfig?.answerJudgment.routeName;
            const genRoute = genRouteName
                ? routesByName.get(genRouteName)
                : undefined;
            const judgeRoute = judgeRouteName
                ? routesByName.get(judgeRouteName)
                : undefined;
            const genProviderName = genRoute?.steps[0]?.providerName;
            const judgeProviderName = judgeRoute?.steps[0]?.providerName;
            const genProvider = genProviderName
                ? providersByName.get(genProviderName)
                : undefined;
            const judgeProvider = judgeProviderName
                ? providersByName.get(judgeProviderName)
                : undefined;

            for (const [rowId, attempts] of Object.entries(rowMap)) {
                if (enabledRowIds && !enabledRowIds.has(rowId)) {
                    continue;
                }
                const row = rowById.get(rowId);

                const questionFields: Record<string, string> = {};
                if (taskConfig && row?.values) {
                    for (const key of taskConfig.answerGeneration
                        .questionFieldKeys) {
                        const title = columnTitleByKey.get(key) ?? key;
                        questionFields[title] = row.values[key]?.value ?? "";
                    }
                }

                const answerFields: Record<string, string> = {};
                if (taskConfig && row?.values) {
                    for (const key of taskConfig.answerJudgment
                        .answerFieldKeys) {
                        const title = columnTitleByKey.get(key) ?? key;
                        answerFields[title] = row.values[key]?.value ?? "";
                    }
                }

                exportItems.push({
                    dataset_name: fileName,
                    source_file_name: state.sourceFileName ?? null,
                    data_source_name: state.dataSourceName ?? null,
                    sample_id: rowId,
                    task_name: taskConfig?.name ?? taskId,

                    question_fields: questionFields,
                    answer_fields: answerFields,

                    generation_route_name: genRouteName ?? null,
                    generation_model: genRoute?.model ?? null,
                    generation_provider: genProvider?.name ?? null,
                    generation_provider_api_type: genProvider?.apiType ?? null,
                    generation_provider_api_url: genProvider?.apiUrl ?? null,
                    judgment_route_name: judgeRouteName ?? null,
                    judgment_model: judgeRoute?.model ?? null,
                    judgment_provider: judgeProvider?.name ?? null,
                    judgment_provider_api_type: judgeProvider?.apiType ?? null,
                    judgment_provider_api_url: judgeProvider?.apiUrl ?? null,

                    generation_system_prompt:
                        taskConfig?.answerGeneration.prompt ?? null,
                    judgment_system_prompt:
                        taskConfig?.answerJudgment.prompt ?? null,

                    attempts: attempts.map((a) => ({
                        attempt_index: a.attemptIndex,
                        updated_at: a.updatedAt ?? null,
                        generation_response_text: a.generationResponseText,
                        generation_parsed_json: safeJsonParse(
                            a.generationParsedJsonText,
                        ),
                        judgment_response_text: a.judgmentResponseText,
                        judgment_parsed_json: safeJsonParse(
                            a.judgmentParsedJsonText,
                        ),
                        final_verdict: a.finalVerdict,
                        generation_latency_ms: a.generationLatencyMs ?? null,
                        judgment_latency_ms: a.judgmentLatencyMs ?? null,
                        generation_input_tokens:
                            a.generationInputTokens ?? null,
                        generation_output_tokens:
                            a.generationOutputTokens ?? null,
                        judgment_input_tokens: a.judgmentInputTokens ?? null,
                        judgment_output_tokens: a.judgmentOutputTokens ?? null,
                        generation_finish_reason:
                            a.generationFinishReason ?? null,
                        judgment_finish_reason: a.judgmentFinishReason ?? null,
                    })),
                });
            }
        }

        const baseName = fileName.replace(/\.[^.]+$/, "");
        const exportName = `${baseName}-评测结果.json`;
        const encodedFileName = encodeURIComponent(exportName);

        res.setHeader("Content-Type", "application/json; charset=utf-8");
        res.setHeader(
            "Content-Disposition",
            `attachment; filename*=UTF-8''${encodedFileName}`,
        );
        return res.send(JSON.stringify(exportItems, null, 2));
    });
};

function safeJsonParse(text: string | undefined): unknown {
    if (!text) return null;
    try {
        return JSON.parse(text);
    } catch {
        return null;
    }
}
