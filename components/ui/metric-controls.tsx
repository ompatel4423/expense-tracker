import React from 'react';
import { LineChart, BarChart2, ChevronDown } from 'lucide-react';
import type { ChartView } from './metric-chart';

export interface PeriodOption {
  label: string;
  points?: number;
}

interface ViewToggleProps {
  value: ChartView;
  onChange: (view: ChartView) => void;
}

export function ViewToggle({ value, onChange }: ViewToggleProps) {
  return (
    <div className="pointer-events-auto inline-flex items-center rounded-full border border-border/80 bg-muted/30 p-0.5">
      <button
        type="button"
        onClick={() => onChange('curve')}
        className={`flex h-7 w-7 items-center justify-center rounded-full transition-all ${
          value === 'curve'
            ? 'bg-background text-foreground shadow-sm'
            : 'text-muted-foreground hover:text-foreground'
        }`}
        title="Curve View"
      >
        <LineChart size={13} strokeWidth={2.5} />
      </button>
      <button
        type="button"
        onClick={() => onChange('bar')}
        className={`flex h-7 w-7 items-center justify-center rounded-full transition-all ${
          value === 'bar'
            ? 'bg-background text-foreground shadow-sm'
            : 'text-muted-foreground hover:text-foreground'
        }`}
        title="Bar View"
      >
        <BarChart2 size={13} strokeWidth={2.5} />
      </button>
    </div>
  );
}

interface PeriodSelectProps {
  value: string;
  options: PeriodOption[];
  onChange: (option: PeriodOption) => void;
  accentText?: string;
}

export function PeriodSelect({ value, options, onChange, accentText }: PeriodSelectProps) {
  return (
    <div className="pointer-events-auto relative inline-flex items-center">
      <select
        value={value}
        onChange={(e) => {
          const selected = options.find((o) => o.label === e.target.value);
          if (selected) onChange(selected);
        }}
        className="appearance-none rounded-full border border-border/80 bg-muted/30 py-1 pl-3 pr-7 text-[12px] font-medium text-foreground transition-all hover:bg-muted/50 focus:outline-none focus:ring-1 focus:ring-ring cursor-pointer"
        style={{ color: accentText }}
      >
        {options.map((opt) => (
          <option key={opt.label} value={opt.label} className="bg-popover text-popover-foreground">
            {opt.label}
          </option>
        ))}
      </select>
      <ChevronDown
        size={12}
        className="pointer-events-none absolute right-2 text-muted-foreground"
      />
    </div>
  );
}
