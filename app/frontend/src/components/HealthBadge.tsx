interface Props {
  label: string;
  meaning: string;
  margin_pct: number;
}

const COLORS: Record<string, string> = {
  STRONG: "bg-emerald-100 text-emerald-800 border-emerald-200",
  "ON TARGET": "bg-blue-100 text-blue-800 border-blue-200",
  REVIEW: "bg-amber-100 text-amber-800 border-amber-200",
  THIN: "bg-amber-100 text-amber-800 border-amber-200",
  "BELOW FLOOR": "bg-red-100 text-red-800 border-red-200",
  "LOSING MONEY": "bg-red-100 text-red-800 border-red-200",
};

function getColorClass(label: string): string {
  for (const [key, cls] of Object.entries(COLORS)) {
    if (label.toUpperCase().includes(key)) return cls;
  }
  return "bg-gray-100 text-gray-800 border-gray-200";
}

export default function HealthBadge({ label, meaning, margin_pct }: Props) {
  return (
    <div className={`rounded-lg border p-3 ${getColorClass(label)}`}>
      <div className="flex items-center justify-between">
        <span className="text-sm font-semibold">{label}</span>
        <span className="text-lg font-bold">{margin_pct.toFixed(1)}%</span>
      </div>
      <p className="mt-1 text-xs opacity-80">{meaning}</p>
    </div>
  );
}
