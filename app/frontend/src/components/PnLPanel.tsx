import HealthBadge from "./HealthBadge";

interface Props {
  summary: any;
  health: any;
}

function fmt(n: number): string {
  return new Intl.NumberFormat("en-US", {
    style: "currency",
    currency: "USD",
  }).format(n);
}

export default function PnLPanel({ summary, health }: Props) {
  if (!summary) return null;

  return (
    <div className="space-y-4">
      {/* Health indicators */}
      {health && (
        <div className="grid grid-cols-2 gap-3">
          <HealthBadge {...health.qc_gross_margin} />
          <HealthBadge {...health.true_net_margin} />
        </div>
      )}

      {/* Financial summary */}
      <div className="card space-y-3">
        <h3 className="text-lg font-semibold text-brand-700">P&L Summary</h3>
        <div className="space-y-2 text-sm">
          <Row label="Total Vendor Cost" value={fmt(summary.total_vendor_cost)} />
          <Row label="Total Client Cost" value={fmt(summary.total_client_cost)} />
          {summary.management_fee > 0 && (
            <Row label="Management Fee" value={fmt(summary.management_fee)} />
          )}
          <div className="border-t border-brand-100 pt-2">
            <Row
              label="Client Invoice Subtotal"
              value={fmt(summary.client_invoice_subtotal)}
              bold
            />
          </div>
          {summary.total_tax > 0 && (
            <Row label="Tax" value={fmt(summary.total_tax)} />
          )}
          <Row
            label="Client Invoice Total"
            value={fmt(summary.client_invoice_total)}
            bold
          />
          <div className="border-t border-brand-100 pt-2">
            <Row
              label={`Commission (${summary.commission_pct}%)`}
              value={`-${fmt(summary.commission_amount)}`}
              muted
            />
            <Row
              label={`CC Fee (${summary.cc_fee_pct}%)`}
              value={`-${fmt(summary.cc_fee_amount)}`}
              muted
            />
          </div>
          <div className="border-t border-brand-100 pt-2">
            <Row
              label="QC Gross Revenue"
              value={fmt(summary.qc_gross_revenue)}
              bold
            />
            <Row
              label="QC Gross Margin"
              value={`${summary.qc_gross_margin_pct.toFixed(1)}%`}
            />
          </div>
          {(summary.opex_cost > 0 || summary.travel_expenses > 0) && (
            <div className="border-t border-brand-100 pt-2">
              {summary.opex_cost > 0 && (
                <Row
                  label={`OpEx (${summary.opex_hours}h × ${fmt(summary.opex_rate)})`}
                  value={`-${fmt(summary.opex_cost)}`}
                  muted
                />
              )}
              {summary.travel_expenses > 0 && (
                <Row
                  label="Travel Expenses"
                  value={`-${fmt(summary.travel_expenses)}`}
                  muted
                />
              )}
            </div>
          )}
          <div className="border-t-2 border-brand-300 pt-2">
            <Row
              label="True Net Profit"
              value={fmt(summary.true_net_profit)}
              bold
              highlight={summary.true_net_profit < 0}
            />
            <Row
              label="True Net Margin"
              value={`${summary.true_net_margin_pct.toFixed(1)}%`}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

function Row({
  label,
  value,
  bold,
  muted,
  highlight,
}: {
  label: string;
  value: string;
  bold?: boolean;
  muted?: boolean;
  highlight?: boolean;
}) {
  return (
    <div className="flex justify-between">
      <span className={muted ? "text-brand-400" : ""}>{label}</span>
      <span
        className={`${bold ? "font-semibold" : ""} ${
          highlight ? "text-red-600 font-bold" : ""
        } ${muted ? "text-brand-400" : ""}`}
      >
        {value}
      </span>
    </div>
  );
}
