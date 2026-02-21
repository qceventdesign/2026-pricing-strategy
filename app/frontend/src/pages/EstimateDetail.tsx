import { useEffect, useState } from "react";
import { useParams, useNavigate, Link } from "react-router-dom";
import { api } from "../api";
import PnLPanel from "../components/PnLPanel";

function fmt(n: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(n);
}

export default function EstimateDetail() {
  const { id } = useParams<{ id: string }>();
  const navigate = useNavigate();
  const [estimate, setEstimate] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    if (!id) return;
    api
      .getEstimate(id)
      .then(setEstimate)
      .catch(() => navigate("/"))
      .finally(() => setLoading(false));
  }, [id, navigate]);

  const handleDelete = async () => {
    if (!id) return;
    if (!confirm("Delete this estimate?")) return;
    await api.deleteEstimate(id);
    navigate("/");
  };

  if (loading) {
    return <div className="py-12 text-center text-brand-400">Loading...</div>;
  }

  if (!estimate) {
    return <div className="py-12 text-center text-brand-400">Not found</div>;
  }

  const calc = estimate.calculated;
  const summary = calc?.summary || calc;
  const health = calc?.health || estimate.health;
  const lineItems = calc?.line_items || [];

  return (
    <div className="space-y-6">
      {/* Header */}
      <div className="flex items-center justify-between">
        <div>
          <Link
            to="/"
            className="text-sm text-brand-400 hover:text-brand-600"
          >
            &larr; Back to Dashboard
          </Link>
          <h1 className="mt-2 text-3xl font-bold text-brand-700">
            {estimate.name}
          </h1>
          <p className="text-brand-400">
            {estimate.client_name && `${estimate.client_name} — `}
            {estimate.event_name}
            {estimate.event_date && ` — ${estimate.event_date}`}
          </p>
        </div>
        <div className="flex gap-2">
          <a
            href={api.exportExcelUrl(id!)}
            className="btn-secondary"
            download
          >
            Export Excel
          </a>
          <button onClick={handleDelete} className="btn-danger">
            Delete
          </button>
        </div>
      </div>

      <div className="grid grid-cols-3 gap-6">
        {/* Left: line items + details (2 cols) */}
        <div className="col-span-2 space-y-6">
          {/* Event info */}
          <div className="card">
            <h2 className="mb-3 text-lg font-semibold text-brand-700">
              Event Details
            </h2>
            <div className="grid grid-cols-3 gap-4 text-sm">
              <Detail label="Client" value={estimate.client_name || "-"} />
              <Detail label="Event" value={estimate.event_name || "-"} />
              <Detail label="Date" value={estimate.event_date || "-"} />
              <Detail
                label="Guests"
                value={estimate.guest_count ? String(estimate.guest_count) : "-"}
              />
              <Detail
                label="Commission"
                value={summary?.commission_label || estimate.commission_scenario}
              />
              <Detail
                label="Type"
                value={estimate.estimate_type}
              />
            </div>
            {estimate.notes && (
              <div className="mt-4 rounded-lg bg-brand-50 p-3 text-sm text-brand-500">
                {estimate.notes}
              </div>
            )}
          </div>

          {/* Line items table */}
          <div className="card">
            <h2 className="mb-3 text-lg font-semibold text-brand-700">
              Line Items
            </h2>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-brand-100">
                    <th className="pb-2 text-left font-medium text-brand-500">
                      Description
                    </th>
                    <th className="pb-2 text-left font-medium text-brand-500">
                      Category
                    </th>
                    <th className="pb-2 text-center font-medium text-brand-500">
                      Qty
                    </th>
                    <th className="pb-2 text-right font-medium text-brand-500">
                      Vendor Cost
                    </th>
                    <th className="pb-2 text-center font-medium text-brand-500">
                      Markup
                    </th>
                    <th className="pb-2 text-right font-medium text-brand-500">
                      Client Cost
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {lineItems.map((item: any, i: number) => (
                    <tr key={i} className="border-b border-brand-50">
                      <td className="py-2">{item.description || "-"}</td>
                      <td className="py-2 text-brand-400">
                        {item.category_label}
                      </td>
                      <td className="py-2 text-center">{item.quantity}</td>
                      <td className="py-2 text-right">
                        {fmt(item.vendor_total)}
                      </td>
                      <td className="py-2 text-center text-brand-400">
                        {item.markup_pct}%
                      </td>
                      <td className="py-2 text-right font-medium">
                        {fmt(item.client_cost)}
                      </td>
                    </tr>
                  ))}
                </tbody>
                {lineItems.length > 0 && (
                  <tfoot>
                    <tr className="border-t-2 border-brand-200">
                      <td colSpan={3} className="py-2 font-semibold">
                        Totals
                      </td>
                      <td className="py-2 text-right font-semibold">
                        {summary && fmt(summary.total_vendor_cost)}
                      </td>
                      <td></td>
                      <td className="py-2 text-right font-semibold">
                        {summary && fmt(summary.total_client_cost)}
                      </td>
                    </tr>
                  </tfoot>
                )}
              </table>
              {lineItems.length === 0 && (
                <p className="py-6 text-center text-sm text-brand-300">
                  No line items
                </p>
              )}
            </div>
          </div>

          {/* Client-ready summary */}
          {summary && (
            <div className="card">
              <h2 className="mb-3 text-lg font-semibold text-brand-700">
                Client Summary
              </h2>
              <div className="grid grid-cols-2 gap-4 text-sm">
                <SummaryRow
                  label="Subtotal"
                  value={fmt(summary.client_invoice_subtotal)}
                />
                {summary.management_fee > 0 && (
                  <SummaryRow
                    label="Management Fee"
                    value={fmt(summary.management_fee)}
                  />
                )}
                {summary.total_tax > 0 && (
                  <SummaryRow label="Tax" value={fmt(summary.total_tax)} />
                )}
                <SummaryRow
                  label="Client Invoice Total"
                  value={fmt(summary.client_invoice_total)}
                  bold
                />
              </div>
            </div>
          )}
        </div>

        {/* Right: P&L panel (1 col) */}
        <div className="col-span-1">
          <div className="sticky top-8">
            <h2 className="mb-4 text-lg font-semibold text-brand-700">
              P&L Analysis
            </h2>
            <PnLPanel summary={summary} health={health} />
          </div>
        </div>
      </div>
    </div>
  );
}

function Detail({ label, value }: { label: string; value: string }) {
  return (
    <div>
      <div className="text-brand-400">{label}</div>
      <div className="font-medium">{value}</div>
    </div>
  );
}

function SummaryRow({
  label,
  value,
  bold,
}: {
  label: string;
  value: string;
  bold?: boolean;
}) {
  return (
    <>
      <div className={bold ? "font-semibold" : "text-brand-500"}>{label}</div>
      <div className={`text-right ${bold ? "font-bold text-lg" : ""}`}>
        {value}
      </div>
    </>
  );
}
