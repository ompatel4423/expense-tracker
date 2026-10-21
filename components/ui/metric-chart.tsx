import React, { useId, useState } from 'react';

export type SeriesPoint = {
  value: number;
  date: string;
};

export type MetricAccent = 'emerald' | 'rose' | 'neutral';

export type MetricSeries = {
  name: string;
  data: SeriesPoint[];
  accent?: MetricAccent;
};

export type ChartSeries = {
  name: string;
  data: SeriesPoint[];
  color: string;
};

export type ChartView = 'curve' | 'bar';

export const ACCENTS: Record<MetricAccent, { stroke: string; text: string; fill: string }> = {
  emerald: { stroke: '#10b981', text: '#34d399', fill: 'rgba(16, 185, 129, 0.15)' },
  rose: { stroke: '#f43f5e', text: '#fb7185', fill: 'rgba(244, 63, 94, 0.15)' },
  neutral: { stroke: '#94a3b8', text: '#94a3b8', fill: 'rgba(148, 163, 184, 0.15)' },
};

export const SERIES_COLORS = ['#38bdf8', '#a855f7', '#f59e0b', '#10b981', '#ec4899'];

export function formatCompact(n: number): string {
  if (Math.abs(n) >= 1_000_000) return (n / 1_000_000).toFixed(1).replace(/\.0$/, '') + 'M';
  if (Math.abs(n) >= 1_000) return (n / 1_000).toFixed(1).replace(/\.0$/, '') + 'k';
  return n.toLocaleString();
}

interface MetricChartProps {
  series: ChartSeries[];
  view: ChartView;
  defaultIndex?: number;
  valueFormatter?: (value: number) => string;
  dateFormatter?: (date: string) => string;
}

function getSplinePath(pts: { x: number; y: number }[], k = 0.2): string {
  if (pts.length === 0) return '';
  let d = `M ${pts[0].x} ${pts[0].y}`;
  for (let i = 0; i < pts.length - 1; i++) {
    const p0 = pts[Math.max(i - 1, 0)];
    const p1 = pts[i];
    const p2 = pts[i + 1];
    const p3 = pts[Math.min(i + 2, pts.length - 1)];

    const cp1x = (p1.x + (p2.x - p0.x) * k).toFixed(2);
    const cp1y = (p1.y + (p2.y - p0.y) * k).toFixed(2);
    const cp2x = (p2.x - (p3.x - p1.x) * k).toFixed(2);
    const cp2y = (p2.y - (p3.y - p1.y) * k).toFixed(2);

    d += ` C ${cp1x} ${cp1y}, ${cp2x} ${cp2y}, ${p2.x} ${p2.y}`;
  }
  return d;
}

export function MetricChart({
  series,
  view,
  defaultIndex,
  valueFormatter,
  dateFormatter,
}: MetricChartProps) {
  const chartId = useId().replace(/:/g, '');
  const [activeIndex, setActiveIndex] = useState<number | null>(defaultIndex ?? null);

  const primary = series[0];
  const points = primary?.data ?? [];
  if (points.length < 2) return null;

  const width = 400;
  const height = 180;
  const paddingX = 20;
  const paddingTop = 25;
  const paddingBottom = 20;

  const usableWidth = width - paddingX * 2;
  const usableHeight = height - paddingTop - paddingBottom;

  const values = points.map((p) => p.value);
  const minVal = Math.min(...values);
  const maxVal = Math.max(...values, minVal + 1);

  const coords = points.map((p, idx) => {
    const x = paddingX + (idx / (points.length - 1)) * usableWidth;
    const ratio = (p.value - minVal) / (maxVal - minVal);
    const y = paddingTop + (1 - ratio) * usableHeight;
    return { x, y, point: p };
  });

  const spline = getSplinePath(coords);
  const area = `${spline} L ${coords[coords.length - 1].x} ${height - paddingBottom} L ${coords[0].x} ${height - paddingBottom} Z`;

  const activePoint = activeIndex !== null && coords[activeIndex] ? coords[activeIndex] : null;

  return (
    <div className="relative h-full w-full select-none">
      <svg
        className="h-full w-full overflow-visible"
        viewBox={`0 0 ${width} ${height}`}
        preserveAspectRatio="none"
      >
        <defs>
          <linearGradient id={`grad-${chartId}`} x1="0" y1="0" x2="0" y2="1">
            <stop offset="0%" stopColor={primary.color} stopOpacity="0.32" />
            <stop offset="80%" stopColor={primary.color} stopOpacity="0.04" />
            <stop offset="100%" stopColor={primary.color} stopOpacity="0" />
          </linearGradient>
        </defs>

        {view === 'curve' ? (
          <>
            {/* Area under curve */}
            <path d={area} fill={`url(#grad-${chartId})`} />
            {/* Main curve line */}
            <path
              d={spline}
              fill="none"
              stroke={primary.color}
              strokeWidth="2.5"
              strokeLinecap="round"
              strokeLinejoin="round"
            />
            {/* Data point dots */}
            {coords.map((c, i) => (
              <g
                key={i}
                className="cursor-pointer"
                onMouseEnter={() => setActiveIndex(i)}
                onClick={() => setActiveIndex(i)}
              >
                <circle
                  cx={c.x}
                  cy={c.y}
                  r={i === activeIndex ? 5.5 : 3.5}
                  fill={primary.color}
                  stroke="#ffffff"
                  strokeWidth={i === activeIndex ? 2 : 1}
                  className="transition-all duration-200"
                />
              </g>
            ))}
          </>
        ) : (
          /* Bar View */
          coords.map((c, i) => {
            const barWidth = Math.max(6, Math.min(18, usableWidth / points.length - 4));
            const barH = height - paddingBottom - c.y;
            const isHovered = i === activeIndex;
            return (
              <g
                key={i}
                className="cursor-pointer"
                onMouseEnter={() => setActiveIndex(i)}
                onClick={() => setActiveIndex(i)}
              >
                {/* Background Track */}
                <rect
                  x={c.x - barWidth / 2}
                  y={paddingTop}
                  width={barWidth}
                  height={usableHeight}
                  rx={barWidth / 2}
                  fill="currentColor"
                  className="text-foreground/[0.04]"
                />
                {/* Active Capsule Fill */}
                <rect
                  x={c.x - barWidth / 2}
                  y={c.y}
                  width={barWidth}
                  height={Math.max(barH, 4)}
                  rx={barWidth / 2}
                  fill={primary.color}
                  opacity={isHovered ? 1 : 0.75}
                  className="transition-all duration-200"
                />
              </g>
            );
          })
        )}
      </svg>

      {/* Floating Tooltip */}
      {activePoint && (
        <div
          className="pointer-events-none absolute z-30 flex -translate-x-1/2 flex-col items-center gap-0.5 rounded-lg border border-border bg-popover/90 px-2.5 py-1 text-xs shadow-md backdrop-blur-md transition-all duration-150"
          style={{
            left: `${(activePoint.x / width) * 100}%`,
            top: `${Math.max(0, (activePoint.y / height) * 100 - 35)}%`,
          }}
        >
          <span className="font-semibold text-popover-foreground">
            {valueFormatter
              ? valueFormatter(activePoint.point.value)
              : activePoint.point.value.toLocaleString()}
          </span>
          <span className="text-[10px] text-muted-foreground">
            {dateFormatter
              ? dateFormatter(activePoint.point.date)
              : activePoint.point.date}
          </span>
        </div>
      )}
    </div>
  );
}
