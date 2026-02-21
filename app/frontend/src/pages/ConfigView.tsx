import { useState, useEffect } from 'react'
import { api } from '../api/client'

export default function ConfigView() {
  const [config, setConfig] = useState<any>(null)
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    api.getConfig().then(setConfig).finally(() => setLoading(false))
  }, [])

  if (loading) return <div className="card text-center py-12 text-gray-500">Loading...</div>
  if (!config) return <div className="card text-center py-12 text-red-500">Failed to load config</div>

  const markups = config.markups
  const rates = config.rates
  const fees = config.fees
  const commissions = config.commissions
  const thresholds = config.thresholds

  return (
    <div>
      <div className="mb-8">
        <h1 className="text-2xl font-bold text-gray-900">Pricing Configuration</h1>
        <p className="text-sm text-gray-500 mt-1">
          Reference view of all pricing rules. Edit the JSON config files to update.
        </p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-2 gap-6">
        {/* Category Markups */}
        <div className="card lg:col-span-2">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Category Markup Schedule</h2>
          <p className="text-sm text-gray-500 mb-4">
            Absolute floor: {markups.absolute_floor_markup_pct}% &middot; Typical blended:{' '}
            {markups.typical_blended_range_pct[0]}–{markups.typical_blended_range_pct[1]}%
          </p>
          <table className="w-full text-sm">
            <thead>
              <tr className="text-xs text-gray-500 uppercase tracking-wider border-b border-gray-200">
                <th className="text-left py-2">Category</th>
                <th className="text-right py-2">Markup</th>
                <th className="text-left py-2 pl-4">Rationale</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(markups.categories).map(([key, cat]: [string, any]) => (
                <tr key={key} className="border-b border-gray-50">
                  <td className="py-2.5 font-medium text-gray-900">{cat.label}</td>
                  <td className="py-2.5 text-right">
                    <span className="font-mono font-bold text-qc-700">{cat.markup_pct}%</span>
                  </td>
                  <td className="py-2.5 pl-4 text-gray-600">{cat.rationale}</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Commission Scenarios */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Commission Scenarios</h2>
          <table className="w-full text-sm">
            <thead>
              <tr className="text-xs text-gray-500 uppercase tracking-wider border-b border-gray-200">
                <th className="text-left py-2">Scenario</th>
                <th className="text-right py-2">Commission</th>
                <th className="text-right py-2">CC Fee</th>
                <th className="text-right py-2">Net to QC</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(commissions.scenarios).map(([key, s]: [string, any]) => (
                <tr key={key} className="border-b border-gray-50">
                  <td className="py-2 font-medium">{s.label}</td>
                  <td className="py-2 text-right">{s.commission_pct}%</td>
                  <td className="py-2 text-right">{s.cc_fee_pct}%</td>
                  <td className="py-2 text-right font-mono font-bold">{s.net_to_qc_pct}%</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Margin Targets */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Margin Targets</h2>
          <table className="w-full text-sm">
            <thead>
              <tr className="text-xs text-gray-500 uppercase tracking-wider border-b border-gray-200">
                <th className="text-left py-2">Metric</th>
                <th className="text-right py-2">Floor</th>
                <th className="text-right py-2">Target</th>
                <th className="text-right py-2">Stretch</th>
              </tr>
            </thead>
            <tbody>
              {Object.entries(thresholds.margin_targets).map(([key, t]: [string, any]) => (
                <tr key={key} className="border-b border-gray-50">
                  <td className="py-2 font-medium">{t.label}</td>
                  <td className="py-2 text-right">{t.floor_pct}%</td>
                  <td className="py-2 text-right">
                    {t.target_range_pct[0]}–{t.target_range_pct[1]}%
                  </td>
                  <td className="py-2 text-right font-bold text-emerald-700">{t.stretch_pct}%+</td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>

        {/* Rates */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Standard Rates</h2>
          <div className="mb-4">
            <h3 className="text-sm font-semibold text-gray-700 mb-2">Tour & Staffing</h3>
            {Object.entries(rates.tour_staffing).map(([key, val]: [string, any]) => (
              <div key={key} className="flex justify-between text-sm py-1">
                <span className="text-gray-600 capitalize">{key.replace(/_/g, ' ')}</span>
                <span className="font-mono">${val}</span>
              </div>
            ))}
          </div>
          <div className="mb-4">
            <h3 className="text-sm font-semibold text-gray-700 mb-2">Transportation</h3>
            {Object.entries(rates.transportation).map(([key, val]: [string, any]) => (
              <div key={key} className="flex justify-between text-sm py-1">
                <span className="text-gray-600 capitalize">{key.replace(/_/g, ' ')}</span>
                <span className="font-mono">${val}</span>
              </div>
            ))}
          </div>
          <div className="text-sm text-gray-500 pt-2 border-t border-gray-100">
            Internal hourly rate: <span className="font-mono font-bold">${rates.internal_hourly_rate}/hr</span>
          </div>
        </div>

        {/* Fee Defaults */}
        <div className="card">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Fee Defaults</h2>
          <div className="space-y-2 text-sm">
            <div className="flex justify-between">
              <span className="text-gray-600">CC Processing Fee</span>
              <span className="font-mono">{fees.cc_processing_pct}%</span>
            </div>
            <div className="flex justify-between">
              <span className="text-gray-600">Service Charge (Restaurants)</span>
              <span className="font-mono">{fees.service_charge_restaurants_pct}%</span>
            </div>
            <div className="flex justify-between">
              <span className="text-gray-600">Guide Gratuity</span>
              <span className="font-mono">{fees.guide_gratuity_pct}%</span>
            </div>
            <div className="flex justify-between">
              <span className="text-gray-600">Admin Fee</span>
              <span className="font-mono">{fees.admin_fee_pct}%</span>
            </div>
            <div className="flex justify-between">
              <span className="text-gray-600">Min Billing Guests (Tours)</span>
              <span className="font-mono">{fees.minimum_billing_guests_tours}</span>
            </div>
            <div className="flex justify-between pt-2 border-t border-gray-100">
              <span className="text-gray-600">Management Fee Range</span>
              <span className="font-mono">
                ${fees.management_fee_range[0].toLocaleString()}–${fees.management_fee_range[1].toLocaleString()}
              </span>
            </div>
            <div className="flex justify-between">
              <span className="text-gray-600">Management Fee Threshold</span>
              <span className="font-mono">${fees.management_fee_threshold.toLocaleString()}</span>
            </div>
          </div>
        </div>

        {/* Health Indicators */}
        <div className="card lg:col-span-2">
          <h2 className="text-lg font-semibold text-gray-900 mb-4">Margin Health Indicators</h2>
          <div className="grid grid-cols-2 gap-6">
            <div>
              <h3 className="text-sm font-semibold text-gray-700 mb-3">QC Gross Margin</h3>
              {Object.entries(thresholds.health_indicators.qc_gross_margin).map(
                ([level, info]: [string, any]) => (
                  <div key={level} className="flex items-center gap-3 py-1.5 text-sm">
                    <span className="font-mono w-20 text-right">{info.threshold_pct >= 0 ? `\u2265${info.threshold_pct}%` : 'Below'}</span>
                    <span className="font-semibold">{info.label}</span>
                    <span className="text-gray-500 text-xs">{info.meaning}</span>
                  </div>
                ),
              )}
            </div>
            <div>
              <h3 className="text-sm font-semibold text-gray-700 mb-3">True Net Margin</h3>
              {Object.entries(thresholds.health_indicators.true_net_margin).map(
                ([level, info]: [string, any]) => (
                  <div key={level} className="flex items-center gap-3 py-1.5 text-sm">
                    <span className="font-mono w-20 text-right">{info.threshold_pct > -100 ? `\u2265${info.threshold_pct}%` : 'Below'}</span>
                    <span className="font-semibold">{info.label}</span>
                    <span className="text-gray-500 text-xs">{info.meaning}</span>
                  </div>
                ),
              )}
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
