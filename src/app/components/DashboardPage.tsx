import { useEffect, useMemo, useState } from "react";
import type {
    FileViewState,
    ParsedColumn,
    StatisticsChartType,
} from "../../types";

interface DashboardPageProps {
    activeFile: FileViewState;
    onOpenStatisticsSettings: () => void;
}

type FieldDistributionItem = {
    label: string;
    count: number;
};

type FieldDistribution = {
    total: number;
    distinctCount: number;
    items: FieldDistributionItem[];
};

const CHART_COLORS = [
    "#0f766e",
    "#0891b2",
    "#ea580c",
    "#7c3aed",
    "#16a34a",
    "#dc2626",
    "#ca8a04",
    "#2563eb",
] as const;

function truncateLabel(value: string, maxLength: number = 18): string {
    return value.length > maxLength ? `${value.slice(0, maxLength)}...` : value;
}

function summarizeItems(
    items: FieldDistributionItem[],
    maxItems: number,
): FieldDistributionItem[] {
    if (items.length <= maxItems) {
        return items;
    }

    const head = items.slice(0, maxItems - 1);
    const otherCount = items
        .slice(maxItems - 1)
        .reduce((sum, item) => sum + item.count, 0);
    return [...head, { label: "其他", count: otherCount }];
}

function BarChart({ items }: { items: FieldDistributionItem[] }) {
    const maxCount = Math.max(...items.map((item) => item.count), 1);
    return (
        <div className="stats-bar-chart">
            {items.map((item, index) => {
                const percent = Math.max(8, (item.count / maxCount) * 100);
                return (
                    <div
                        key={`${item.label}-${index}`}
                        className="stats-bar-row"
                    >
                        <div className="stats-bar-label" title={item.label}>
                            {truncateLabel(item.label)}
                        </div>
                        <div className="stats-bar-track">
                            <div
                                className="stats-bar-fill"
                                style={{
                                    width: `${percent}%`,
                                    background:
                                        CHART_COLORS[
                                            index % CHART_COLORS.length
                                        ],
                                }}
                            />
                        </div>
                        <strong>{item.count}</strong>
                    </div>
                );
            })}
        </div>
    );
}

function PieChart({ items }: { items: FieldDistributionItem[] }) {
    const total = Math.max(
        items.reduce((sum, item) => sum + item.count, 0),
        1,
    );
    let offset = 0;
    const gradientStops = items
        .map((item, index) => {
            const start = (offset / total) * 100;
            offset += item.count;
            const end = (offset / total) * 100;
            return `${CHART_COLORS[index % CHART_COLORS.length]} ${start}% ${end}%`;
        })
        .join(", ");

    return (
        <div className="stats-pie-layout">
            <div
                className="stats-pie-chart"
                style={{
                    background: `conic-gradient(${gradientStops})`,
                }}
            >
                <div className="stats-pie-hole">
                    <strong>{total}</strong>
                    <span>总计</span>
                </div>
            </div>
            <div className="stats-legend">
                {items.map((item, index) => {
                    const percent = ((item.count / total) * 100).toFixed(1);
                    return (
                        <div
                            key={`${item.label}-${index}`}
                            className="stats-legend-item"
                        >
                            <span
                                className="stats-legend-dot"
                                style={{
                                    background:
                                        CHART_COLORS[
                                            index % CHART_COLORS.length
                                        ],
                                }}
                            />
                            <span title={item.label}>
                                {truncateLabel(item.label, 16)}
                            </span>
                            <strong>{`${percent}%`}</strong>
                        </div>
                    );
                })}
            </div>
        </div>
    );
}

function LineChart({ items }: { items: FieldDistributionItem[] }) {
    const width = 320;
    const height = 180;
    const maxCount = Math.max(...items.map((item) => item.count), 1);
    const xStep = items.length > 1 ? 260 / (items.length - 1) : 0;
    const points = items.map((item, index) => {
        const x = 28 + index * xStep;
        const y = 146 - (item.count / maxCount) * 92;
        return { ...item, x, y };
    });
    const polyline = points.map((point) => `${point.x},${point.y}`).join(" ");

    return (
        <div className="stats-line-chart-wrap">
            <svg
                className="stats-line-chart"
                viewBox={`0 0 ${width} ${height}`}
                role="img"
                aria-label="统计折线图"
            >
                <line
                    x1="28"
                    y1="146"
                    x2="292"
                    y2="146"
                    className="stats-axis"
                />
                <line x1="28" y1="24" x2="28" y2="146" className="stats-axis" />
                <polyline points={polyline} className="stats-line" />
                {points.map((point, index) => (
                    <g key={`${point.label}-${index}`}>
                        <circle
                            cx={point.x}
                            cy={point.y}
                            r="4"
                            className="stats-point"
                        />
                        <text
                            x={point.x}
                            y={point.y - 10}
                            className="stats-point-value"
                        >
                            {point.count}
                        </text>
                        <text
                            x={point.x}
                            y="166"
                            textAnchor="middle"
                            className="stats-axis-label"
                        >
                            {truncateLabel(point.label, 8)}
                        </text>
                    </g>
                ))}
            </svg>
        </div>
    );
}

function TableChart({ items }: { items: FieldDistributionItem[] }) {
    return (
        <div className="stats-table-wrap">
            <table className="stats-table">
                <thead>
                    <tr>
                        <th>分类</th>
                        <th>数量</th>
                    </tr>
                </thead>
                <tbody>
                    {items.map((item, index) => (
                        <tr key={`${item.label}-${index}`}>
                            <td title={item.label}>{item.label}</td>
                            <td>{item.count}</td>
                        </tr>
                    ))}
                </tbody>
            </table>
        </div>
    );
}

function FieldChart({
    chartType,
    distribution,
}: {
    chartType: StatisticsChartType;
    distribution: FieldDistribution;
}) {
    const items =
        chartType === "table"
            ? distribution.items
            : summarizeItems(distribution.items, chartType === "pie" ? 6 : 8);

    if (items.length === 0) {
        return (
            <div className="stats-chart-empty">当前字段还没有可统计的数据</div>
        );
    }

    if (chartType === "pie") {
        return <PieChart items={items} />;
    }
    if (chartType === "line") {
        return <LineChart items={items} />;
    }
    if (chartType === "table") {
        return <TableChart items={items} />;
    }
    return <BarChart items={items} />;
}

function getColumnTitle(columns: ParsedColumn[], fieldKey: string): string {
    return columns.find((column) => column.key === fieldKey)?.title ?? fieldKey;
}

export function DashboardPage({
    activeFile,
    onOpenStatisticsSettings,
}: DashboardPageProps) {
    const selectedFieldKeys = activeFile.statisticsConfig.selectedFieldKeys;
    const [rowCount, setRowCount] = useState(activeFile.rowCount ?? 0);
    const [distributionMap, setDistributionMap] = useState<
        Record<string, FieldDistribution>
    >({});

    useEffect(() => {
        setRowCount(activeFile.rowCount ?? 0);
        if (selectedFieldKeys.length === 0) {
            setDistributionMap({});
            return;
        }

        const controller = new AbortController();
        const loadStatistics = async () => {
            try {
                const params = new URLSearchParams({
                    fieldKeys: JSON.stringify(selectedFieldKeys),
                });
                const response = await fetch(
                    `/api/files/${encodeURIComponent(activeFile.fileId)}/statistics?${params.toString()}`,
                    { signal: controller.signal },
                );
                if (!response.ok) {
                    return;
                }
                const payload = (await response.json()) as {
                    rowCount?: unknown;
                    distributions?: unknown;
                };
                if (typeof payload.rowCount === "number") {
                    setRowCount(payload.rowCount);
                }
                const rawDistributions =
                    payload.distributions &&
                    typeof payload.distributions === "object"
                        ? (payload.distributions as Record<string, unknown>)
                        : {};
                const nextMap = Object.entries(rawDistributions).reduce<
                    Record<string, FieldDistribution>
                >((acc, [fieldKey, value]) => {
                    if (!value || typeof value !== "object") {
                        return acc;
                    }
                    const candidate = value as {
                        total?: unknown;
                        distinctCount?: unknown;
                        items?: unknown;
                    };
                    acc[fieldKey] = {
                        total:
                            typeof candidate.total === "number"
                                ? candidate.total
                                : 0,
                        distinctCount:
                            typeof candidate.distinctCount === "number"
                                ? candidate.distinctCount
                                : 0,
                        items: Array.isArray(candidate.items)
                            ? candidate.items.filter(
                                  (item): item is FieldDistributionItem =>
                                      !!item &&
                                      typeof item === "object" &&
                                      typeof (
                                          item as FieldDistributionItem
                                      ).label === "string" &&
                                      typeof (
                                          item as FieldDistributionItem
                                      ).count === "number",
                              )
                            : [],
                    };
                    return acc;
                }, {});
                setDistributionMap(nextMap);
            } catch {
                // Ignore aborted / network errors.
            }
        };

        void loadStatistics();
        return () => controller.abort();
    }, [activeFile.fileId, activeFile.rowCount, selectedFieldKeys]);

    const fieldCards = useMemo(
        () =>
            selectedFieldKeys.map((fieldKey) => {
                const title = getColumnTitle(activeFile.columns, fieldKey);
                const chartType =
                    activeFile.statisticsConfig.chartTypeByField[fieldKey] ??
                    "bar";
                return {
                    fieldKey,
                    title,
                    chartType,
                    distribution: distributionMap[fieldKey] ?? {
                        total: 0,
                        distinctCount: 0,
                        items: [],
                    },
                };
            }),
        [activeFile, distributionMap, selectedFieldKeys],
    );

    const totalDistinctValues = fieldCards.reduce(
        (sum, field) => sum + field.distribution.distinctCount,
        0,
    );

    return (
        <div className="dashboard-page">
            <section className="dashboard-hero">
                <div className="dashboard-hero-copy">
                    <span className="dashboard-eyebrow">统计主页</span>
                    <h2>{activeFile.fileName}</h2>
                    <p>
                        按数据源查看分类分布，统计项和图表类型可在设置页单独配置。
                    </p>
                </div>
                <div className="dashboard-hero-actions">
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onOpenStatisticsSettings}
                    >
                        统计设置
                    </button>
                </div>
            </section>

            <section className="dashboard-summary-grid">
                <article className="dashboard-summary-card">
                    <span>题目总数</span>
                    <strong>{rowCount}</strong>
                    <p>当前数据源下的全部记录数</p>
                </article>
                <article className="dashboard-summary-card">
                    <span>统计字段</span>
                    <strong>{fieldCards.length}</strong>
                    <p>已启用用于展示的分类字段</p>
                </article>
                <article className="dashboard-summary-card">
                    <span>字段维度</span>
                    <strong>{totalDistinctValues}</strong>
                    <p>已启用字段的去重分类总量</p>
                </article>
                <article className="dashboard-summary-card">
                    <span>总字段数</span>
                    <strong>{activeFile.columns.length}</strong>
                    <p>当前数据源可选统计字段</p>
                </article>
            </section>

            {fieldCards.length === 0 ? (
                <section className="dashboard-empty-state">
                    <h3>还没有配置统计字段</h3>
                    <p>
                        去设置页勾选需要展示的字段，再为每个字段指定图表类型。
                    </p>
                    <button
                        type="button"
                        className="btn btn-primary"
                        onClick={onOpenStatisticsSettings}
                    >
                        去配置统计字段
                    </button>
                </section>
            ) : (
                <section className="dashboard-chart-grid">
                    {fieldCards.map((field) => (
                        <article
                            key={field.fieldKey}
                            className="dashboard-chart-card"
                        >
                            <header className="dashboard-chart-head">
                                <div>
                                    <h3>{field.title}</h3>
                                    <p>{`共 ${field.distribution.distinctCount} 个分类，累计 ${field.distribution.total} 条统计值`}</p>
                                </div>
                                <span className="settings-tag">
                                    {field.chartType === "bar"
                                        ? "柱状图"
                                        : field.chartType === "pie"
                                          ? "饼图"
                                          : field.chartType === "line"
                                            ? "折线图"
                                            : "表格"}
                                </span>
                            </header>
                            <FieldChart
                                chartType={field.chartType}
                                distribution={field.distribution}
                            />
                        </article>
                    ))}
                </section>
            )}
        </div>
    );
}
