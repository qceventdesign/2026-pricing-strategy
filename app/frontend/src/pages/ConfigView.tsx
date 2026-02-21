import { useEffect, useState } from "react";
import { api } from "../api";

export default function ConfigView() {
  const [config, setConfig] = useState<any>(null);
  const [loading, setLoading] = useState(true);

  useEffect(() => {
    api
      .getConfig()
      .then(setConfig)
      .finally(() => setLoading(false));
  }, []);

  if (loading) {
    return <div className="py-12 text-center text-brand-400">Loading...</div>;
  }

  if (!config) {
    return (
      <div className="py-12 text-center text-brand-400">
        Failed to load config
      </div>
    );
  }

  const markups = config.markups;
  const rates = config.rates;
  const fees = config.fees;
  const commissions = config.commissions;
  const thresholds = config.thresholds;

  return (
    <div className="space-y-8">
      <div>
        <h1 className="text-3xl font-bold text-brand-700">Pricing Rules</h1>
        <p className="mt-1 text-brand-400">
          2026 pricing configuration — all rules from config files
        </p>
      </div>

      {/* Category Markups */}
      <div className="card">
        <h2 className="mb-4 text-xl font-semibold text-brand-700">
          Category Markup Schedule
        </h2>
        <p className="mb-3 text-sm text-brand-400">
          Absolute floor: {markups.absolute_floor_markup_pct}% | Typical blended
          range: {markups.typical_blended_range_pct[0]}–
          {markups.typical_blended_range_pct[1]}%
        </p>
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-brand-100">
              <th className="pb-2 text-left font-medium text-brand-500">
                Category
              </th>
              <th className="pb-2 text-right font-medium text-brand-500">
                Markup %
              </th>
              <th className="pb-2 text-left font-medium text-brand-500 pl-6">
                Rationale
              </th>
            </tr>
          </thead>
          <tbody>
            {Object.entries(markups.categories).map(
              ([key, cat]: [string, any]) => (
                <tr key={key} className="border-b border-brand-50">
                  <td className="py-2 font-medium">{cat.label}</td>
                  <td className="py-2 text-right">
                    <span
                      className={`inline-block rounded-full px-2 py-0.5 text-xs font-semibold ${
                        cat.markup_pct >= 85
                          ? "bg-emerald-100 text-emerald-700"
                          : cat.markup_pct >= 65
                          ? "bg-blue-100 text-blue-700"
                          : "bg-brand-100 text-brand-600"
                      }`}
                    >
                      {cat.markup_pct}%
                    </span>
                  </td>
                  <td className="py-2 pl-6 text-brand-400">
                    {cat.rationale}
                  </td>
                </tr>
              )
            )}
          </tbody>
        </table>
      </div>

      {/* Commission Scenarios */}
      <div className="card">
        <h2 className="mb-4 text-xl font-semibold text-brand-700">
          Commission Scenarios
        </h2>
        <div className="grid grid-cols-3 gap-4">
          {Object.entries(commissions.scenarios).map(
            ([key, sc]: [string, any]) => (
              <div
                key={key}
                className="rounded-lg border border-brand-100 p-4"
              >
                <h3 className="font-semibold text-brand-700">{sc.label}</h3>
                <div className="mt-2 space-y-1 text-sm">
                  <div className="flex justify-between">
                    <span className="text-brand-400">Commission</span>
                    <span className="font-medium">{sc.commission_pct}%</span>
                  </div>
                  <div className="flex justify-between">
                    <span className="text-brand-400">CC Fee</span>
                    <span className="font-medium">{sc.cc_fee_pct}%</span>
                  </div>
                  <div className="flex justify-between border-t border-brand-50 pt-1">
                    <span className="text-brand-400">Net to QC</span>
                    <span className="font-semibold">{sc.net_to_qc_pct}%</span>
                  </div>
                </div>
              </div>
            )
          )}
        </div>
      </div>

      {/* Rates */}
      <div className="grid grid-cols-2 gap-6">
        <div className="card">
          <h2 className="mb-4 text-xl font-semibold text-brand-700">
            Staffing Rates
          </h2>
          <div className="space-y-2 text-sm">
            {Object.entries(rates.tour_staffing).map(
              ([key, val]: [string, any]) => (
                <div key={key} className="flex justify-between">
                  <span className="text-brand-400">
                    {key.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase())}
                  </span>
                  <span className="font-medium">${val}</span>
                </div>
              )
            )}
          </div>
        </div>

        <div className="card">
          <h2 className="mb-4 text-xl font-semibold text-brand-700">
            Transportation Rates
          </h2>
          <div className="space-y-2 text-sm">
            {Object.entries(rates.transportation).map(
              ([key, val]: [string, any]) => (
                <div key={key} className="flex justify-between">
                  <span className="text-brand-400">
                    {key.replace(/_/g, " ").replace(/\b\w/g, (c) => c.toUpperCase())}
                  </span>
                  <span className="font-medium">${val}</span>
                </div>
              )
            )}
          </div>
        </div>
      </div>

      {/* OpEx Tiers */}
      <div className="card">
        <h2 className="mb-4 text-xl font-semibold text-brand-700">
          OpEx Tiers
        </h2>
        <p className="mb-3 text-sm text-brand-400">
          Internal hourly rate: ${rates.internal_hourly_rate}/hr
        </p>
        <div className="grid grid-cols-4 gap-4">
          {Object.entries(rates.opex_tiers).map(
            ([key, tier]: [string, any]) => (
              <div
                key={key}
                className="rounded-lg border border-brand-100 p-4"
              >
                <h3 className="font-semibold capitalize text-brand-700">
                  {key}
                </h3>
                <p className="mt-1 text-xs text-brand-400">
                  {tier.description}
                </p>
                <div className="mt-2 text-sm">
                  <div>
                    {tier.hours_range[0]}–{tier.hours_range[1]} hrs
                  </div>
                  <div className="font-medium">
                    ${tier.cost_range[0].toLocaleString()}–$
                    {tier.cost_range[1].toLocaleString()}
                  </div>
                </div>
              </div>
            )
          )}
        </div>
      </div>

      {/* Fees */}
      <div className="card">
        <h2 className="mb-4 text-xl font-semibold text-brand-700">
          Fee Schedule
        </h2>
        <div className="grid grid-cols-2 gap-4 text-sm">
          <div className="flex justify-between">
            <span className="text-brand-400">CC Processing</span>
            <span className="font-medium">{fees.cc_processing_pct}%</span>
          </div>
          <div className="flex justify-between">
            <span className="text-brand-400">Service Charge (Restaurants)</span>
            <span className="font-medium">
              {fees.service_charge_restaurants_pct}%
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-brand-400">Guide Gratuity</span>
            <span className="font-medium">{fees.guide_gratuity_pct}%</span>
          </div>
          <div className="flex justify-between">
            <span className="text-brand-400">Admin Fee</span>
            <span className="font-medium">{fees.admin_fee_pct}%</span>
          </div>
          <div className="flex justify-between">
            <span className="text-brand-400">Min Billing Guests (Tours)</span>
            <span className="font-medium">
              {fees.minimum_billing_guests_tours}
            </span>
          </div>
          <div className="flex justify-between">
            <span className="text-brand-400">Management Fee Threshold</span>
            <span className="font-medium">
              ${fees.management_fee_threshold.toLocaleString()}
            </span>
          </div>
        </div>
      </div>

      {/* Margin Targets */}
      <div className="card">
        <h2 className="mb-4 text-xl font-semibold text-brand-700">
          Margin Targets & Health Indicators
        </h2>
        <table className="w-full text-sm">
          <thead>
            <tr className="border-b border-brand-100">
              <th className="pb-2 text-left font-medium text-brand-500">
                Scenario
              </th>
              <th className="pb-2 text-center font-medium text-brand-500">
                Floor
              </th>
              <th className="pb-2 text-center font-medium text-brand-500">
                Target Range
              </th>
              <th className="pb-2 text-center font-medium text-brand-500">
                Stretch
              </th>
            </tr>
          </thead>
          <tbody>
            {Object.entries(thresholds.margin_targets).map(
              ([key, target]: [string, any]) => (
                <tr key={key} className="border-b border-brand-50">
                  <td className="py-2 font-medium">{target.label}</td>
                  <td className="py-2 text-center">{target.floor_pct}%</td>
                  <td className="py-2 text-center">
                    {target.target_range_pct[0]}–{target.target_range_pct[1]}%
                  </td>
                  <td className="py-2 text-center">{target.stretch_pct}%</td>
                </tr>
              )
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
