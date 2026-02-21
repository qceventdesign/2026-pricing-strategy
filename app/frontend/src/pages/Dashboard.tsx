import { useEffect, useState } from "react";
import { Link } from "react-router-dom";
import { api } from "../api";

const TYPE_LABELS: Record<string, string> = {
  venue: "Venue",
  decor: "Decor",
  transportation: "Transportation",
  entertainment: "Entertainment",
  tour: "Tour",
};

function fmt(n: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(n);
}

export default function Dashboard() {
  const [estimates, setEstimates] = useState<any[]>([]);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    api.listEstimates().then((data) => {
      setEstimates(data);
      setLoading(false);
    }).catch(() => setLoading(false));
  }, []);

  // Aggregate stats
  const totalRevenue = estimates.reduce(
    (sum, e) => sum + (e.calculated?.client_invoice_subtotal || 0),
    0
  );
  const totalProfit = estimates.reduce(
    (sum, e) => sum + (e.calculated?.true_net_profit || 0),
    0
  );
  const avgMargin =
    estimates.length > 0
      ? estimates.reduce(
          (sum, e) => sum + (e.calculated?.qc_gross_margin_pct || 0),
          0
        ) / estimates.length
      : 0;

  return (
    <div className="space-y-8">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-3xl font-bold text-brand-700">Dashboard</h1>
          <p className="mt-1 text-brand-400">
            Manage estimates and track margins
          </p>
        </div>
        {/* New estimate dropdown */}
        <div className="relative group">
          <button className="btn-primary">+ New Estimate</button>
          <div className="invisible group-hover:visible absolute right-0 top-full mt-1 w-48 rounded-lg border border-brand-100 bg-white py-1 shadow-lg z-10">
            {Object.entries(TYPE_LABELS).map(([key, label]) => (
              <Link
                key={key}
                to={`/new/${key}`}
                className="block px-4 py-2 text-sm text-brand-600 hover:bg-brand-50"
              >
                {label} Estimate
              </Link>
            ))}
          </div>
        </div>
      </div>

      {/* Stats cards */}
      {estimates.length > 0 && (
        <div className="grid grid-cols-3 gap-4">
          <div className="card">
            <div className="text-sm text-brand-400">Total Client Revenue</div>
            <div className="mt-1 text-2xl font-bold text-brand-700">
              {fmt(totalRevenue)}
            </div>
          </div>
          <div className="card">
            <div className="text-sm text-brand-400">Total Net Profit</div>
            <div
              className={`mt-1 text-2xl font-bold ${
                totalProfit >= 0 ? "text-emerald-600" : "text-red-600"
              }`}
            >
              {fmt(totalProfit)}
            </div>
          </div>
          <div className="card">
            <div className="text-sm text-brand-400">Avg Gross Margin</div>
            <div className="mt-1 text-2xl font-bold text-brand-700">
              {avgMargin.toFixed(1)}%
            </div>
          </div>
        </div>
      )}

      {/* Estimates table */}
      {loading ? (
        <div className="py-12 text-center text-brand-400">Loading...</div>
      ) : estimates.length === 0 ? (
        <div className="card py-12 text-center">
          <p className="text-brand-400">No estimates yet.</p>
          <p className="mt-2 text-sm text-brand-300">
            Create your first estimate using the button above.
          </p>
        </div>
      ) : (
        <div className="card overflow-hidden p-0">
          <table className="w-full text-sm">
            <thead>
              <tr className="border-b border-brand-100 bg-brand-50/50">
                <th className="px-4 py-3 text-left font-medium text-brand-500">
                  Name
                </th>
                <th className="px-4 py-3 text-left font-medium text-brand-500">
                  Type
                </th>
                <th className="px-4 py-3 text-left font-medium text-brand-500">
                  Client
                </th>
                <th className="px-4 py-3 text-right font-medium text-brand-500">
                  Client Total
                </th>
                <th className="px-4 py-3 text-right font-medium text-brand-500">
                  Gross Margin
                </th>
                <th className="px-4 py-3 text-right font-medium text-brand-500">
                  Net Profit
                </th>
                <th className="px-4 py-3 text-center font-medium text-brand-500">
                  Health
                </th>
              </tr>
            </thead>
            <tbody>
              {estimates.map((est) => (
                <tr
                  key={est.id}
                  className="border-b border-brand-50 hover:bg-brand-50/30"
                >
                  <td className="px-4 py-3">
                    <Link
                      to={`/estimate/${est.id}`}
                      className="font-medium text-brand-700 hover:underline"
                    >
                      {est.name}
                    </Link>
                  </td>
                  <td className="px-4 py-3 text-brand-400">
                    {TYPE_LABELS[est.estimate_type] || est.estimate_type}
                  </td>
                  <td className="px-4 py-3 text-brand-400">
                    {est.client_name || "-"}
                  </td>
                  <td className="px-4 py-3 text-right">
                    {est.calculated
                      ? fmt(est.calculated.client_invoice_subtotal)
                      : "-"}
                  </td>
                  <td className="px-4 py-3 text-right">
                    {est.calculated
                      ? `${est.calculated.qc_gross_margin_pct.toFixed(1)}%`
                      : "-"}
                  </td>
                  <td
                    className={`px-4 py-3 text-right font-medium ${
                      est.calculated?.true_net_profit < 0
                        ? "text-red-600"
                        : "text-emerald-600"
                    }`}
                  >
                    {est.calculated
                      ? fmt(est.calculated.true_net_profit)
                      : "-"}
                  </td>
                  <td className="px-4 py-3 text-center">
                    {est.health?.qc_gross_margin && (
                      <HealthPill label={est.health.qc_gross_margin.label} />
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  );
}

function HealthPill({ label }: { label: string }) {
  const upper = label.toUpperCase();
  let cls = "bg-gray-100 text-gray-600";
  if (upper.includes("STRONG")) cls = "bg-emerald-100 text-emerald-700";
  else if (upper.includes("TARGET")) cls = "bg-blue-100 text-blue-700";
  else if (upper.includes("REVIEW") || upper.includes("THIN"))
    cls = "bg-amber-100 text-amber-700";
  else if (upper.includes("BELOW") || upper.includes("LOSING"))
    cls = "bg-red-100 text-red-700";

  return (
    <span className={`inline-block rounded-full px-2 py-0.5 text-xs font-medium ${cls}`}>
      {label}
    </span>
  );
}
