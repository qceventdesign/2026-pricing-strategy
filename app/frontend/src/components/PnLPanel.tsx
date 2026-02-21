import type { PnL } from '../types'
import HealthBadge from './HealthBadge'

function fmt(n: number): string {
  return n.toLocaleString('en-US', { style: 'currency', currency: 'USD' })
}

function pct(n: number): string {
  return `${n.toFixed(1)}%`
}

export default function PnLPanel({ pnl }: { pnl: PnL }) {
  return (
    <div className="card">
      <h3 className="text-lg font-semibold text-gray-900 mb-4">Profit & Loss Analysis</h3>

      <div className="space-y-3">
        <Row label="Client Invoice Total" value={fmt(pnl.client_invoice_total)} bold />
        <Row label="Commission" value={`- ${fmt(pnl.commission_amount)}`} muted />
        <Row label="GDP Fee" value={`- ${fmt(pnl.gdp_fee)}`} muted />
        <Row label="Total Vendor Costs" value={`- ${fmt(pnl.total_vendor_costs)}`} muted />

        <div className="border-t border-gray-200 pt-3">
          <div className="flex items-center justify-between">
            <div>
              <span className="font-semibold text-gray-900">QC Gross Revenue</span>
              <span className="ml-2 text-sm text-gray-500">{pct(pnl.qc_margin_pct)} margin</span>
            </div>
            <div className="flex items-center gap-3">
              <span className="font-bold text-gray-900">{fmt(pnl.qc_gross_revenue)}</span>
              <HealthBadge health={pnl.qc_margin_health} />
            </div>
          </div>
        </div>

        <Row label={`OpEx (${pnl.opex_hours}h x $${pnl.opex_rate}/hr)`} value={`- ${fmt(pnl.opex_cost)}`} muted />

        <div className="border-t-2 border-gray-900 pt-3">
          <div className="flex items-center justify-between">
            <div>
              <span className="font-bold text-gray-900">True Net Profit</span>
              <span className="ml-2 text-sm text-gray-500">{pct(pnl.true_net_margin_pct)} margin</span>
            </div>
            <div className="flex items-center gap-3">
              <span className={`font-bold text-lg ${pnl.true_net_profit >= 0 ? 'text-emerald-700' : 'text-red-700'}`}>
                {fmt(pnl.true_net_profit)}
              </span>
              <HealthBadge health={pnl.true_net_health} />
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}

function Row({ label, value, bold, muted }: { label: string; value: string; bold?: boolean; muted?: boolean }) {
  return (
    <div className="flex items-center justify-between">
      <span className={`text-sm ${bold ? 'font-semibold text-gray-900' : 'text-gray-600'}`}>{label}</span>
      <span className={`text-sm ${bold ? 'font-bold' : ''} ${muted ? 'text-gray-500' : 'text-gray-900'}`}>
        {value}
      </span>
    </div>
  )
}
