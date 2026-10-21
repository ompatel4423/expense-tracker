import ProgressMetricCard, { type SeriesPoint } from './progress-metric-card';

const ordersSeries: SeriesPoint[] = [
  { value: 100, date: '18 Mar, 2025' },
  { value: 115, date: '21 Mar, 2025' },
  { value: 120, date: '24 Mar, 2025' },
  { value: 110, date: '27 Mar, 2025' },
  { value: 178, date: '30 Mar, 2025' },
  { value: 180, date: '02 Apr, 2025' },
  { value: 205, date: '05 Apr, 2025' },
  { value: 185, date: '08 Apr, 2025' },
  { value: 180, date: '11 Apr, 2025' },
  { value: 183, date: '14 Apr, 2025' },
  { value: 185, date: '17 Apr, 2025' },
  { value: 185, date: '20 Apr, 2025' },
  { value: 166, date: '23 Apr, 2025' },
  { value: 181, date: '26 Apr, 2025' },
  { value: 168, date: '29 Apr, 2025' },
  { value: 195, date: '02 May, 2025' },
  { value: 188, date: '05 May, 2025' },
  { value: 162, date: '08 May, 2025' },
  { value: 166, date: '11 May, 2025' },
];

export default function ProgressMetricCardDemo() {
  return (
    <div className="flex w-full max-w-[760px] flex-col gap-4 p-12">
      <ProgressMetricCard title="Total orders" unit="orders" data={ordersSeries} defaultIndex={6} />
    </div>
  );
}
