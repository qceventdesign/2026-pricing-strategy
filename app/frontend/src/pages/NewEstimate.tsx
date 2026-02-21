import { useState, useEffect } from 'react'
import { useParams, useNavigate } from 'react-router-dom'
import { api } from '../api/client'
import type { LineItem, EstimateTemplate, Category, CalculationResult } from '../types'
import PnLPanel from '../components/PnLPanel'

const COMMISSION_OPTIONS = [
  { key: 'direct_client', label: 'Direct Client (0%)' },
  { key: 'gdp_standard', label: 'GDP Standard (6.5%)' },
  { key: 'gdp_plus_client', label: 'GDP + Client (21.5%)' },
]

function fmt(n: number): string {
  return n.toLocaleString('en-US', { style: 'currency', currency: 'USD' })
}

export default function NewEstimate() {
  const { type } = useParams<{ type: string }>()
  const navigate = useNavigate()

  const [template, setTemplate] = useState<EstimateTemplate | null>(null)
  const [categories, setCategories] = useState<Category[]>([])
  const [saving, setSaving] = useState(false)
  const [result, setResult] = useState<CalculationResult | null>(null)

  // Form state
  const [form, setForm] = useState({
    estimate_name: '',
    client_name: '',
    event_name: '',
    event_date: '',
    event_time: '',
    guest_count: 0,
    location: '',
    tax_rate: 0.07,
    commission_scenario: 'direct_client',
    cc_processing_pct: 0.035,
    gratuity_pct: 0.20,
    admin_fee_pct: 0.05,
    gdp_enabled: false,
    opex_hours: 25,
  })

  const [lineItems, setLineItems] = useState<LineItem[]>([])

  useEffect(() => {
    if (!type) return
    Promise.all([api.getTemplate(type), api.getCategories()]).then(([tmpl, cats]) => {
      setTemplate(tmpl)
      setCategories(cats)
      // Initialize line items from template
      const items: LineItem[] = []
      let order = 0
      for (const section of tmpl.sections) {
        for (const item of section.default_items) {
          items.push({
            section: section.name,
            name: item.name,
            category_key: section.category_key,
            item_type: item.item_type as LineItem['item_type'],
            unit_price: 0,
            quantity: item.item_type === 'flat_fee' ? 1 : 0,
            is_taxable: section.is_taxable,
            notes: '',
            sort_order: order++,
          })
        }
      }
      setLineItems(items)
      // Set GDP enabled based on commission scenario
      setForm((f) => ({
        ...f,
        gdp_enabled: false,
      }))
    })
  }, [type])

  const updateLineItem = (index: number, field: keyof LineItem, value: any) => {
    setLineItems((prev) => {
      const updated = [...prev]
      updated[index] = { ...updated[index], [field]: value }
      return updated
    })
    setResult(null) // Clear previous calculation
  }

  const addLineItem = (sectionName: string, categoryKey: string, isTaxable: boolean) => {
    setLineItems((prev) => [
      ...prev,
      {
        section: sectionName,
        name: '',
        category_key: categoryKey,
        item_type: 'quantity',
        unit_price: 0,
        quantity: 0,
        is_taxable: isTaxable,
        notes: '',
        sort_order: prev.length,
      },
    ])
  }

  const removeLineItem = (index: number) => {
    setLineItems((prev) => prev.filter((_, i) => i !== index))
    setResult(null)
  }

  const handleCalculate = async () => {
    setSaving(true)
    try {
      const data = {
        ...form,
        estimate_type: type!,
        gdp_enabled: form.commission_scenario !== 'direct_client',
        line_items: lineItems,
      }
      const res = await api.createEstimate(data)
      setResult(res)
    } catch (err: any) {
      alert(`Error: ${err.message}`)
    } finally {
      setSaving(false)
    }
  }

  const handleSaveAndView = async () => {
    if (result?.id) {
      navigate(`/estimate/${result.id}`)
    } else {
      await handleCalculate()
    }
  }

  if (!template) {
    return <div className="card text-center py-12 text-gray-500">Loading template...</div>
  }

  // Group line items by section
  const sections = new Map<string, { items: { item: LineItem; globalIndex: number }[]; categoryKey: string; isTaxable: boolean }>()
  lineItems.forEach((item, index) => {
    if (!sections.has(item.section)) {
      sections.set(item.section, { items: [], categoryKey: item.category_key, isTaxable: item.is_taxable })
    }
    sections.get(item.section)!.items.push({ item, globalIndex: index })
  })

  const categoryLabel = (key: string) => categories.find((c) => c.key === key)?.label || key
  const categoryMarkup = (key: string) => categories.find((c) => c.key === key)?.markup_pct || 0

  return (
    <div>
      <div className="mb-6">
        <h1 className="text-2xl font-bold text-gray-900">{template.label}</h1>
        <p className="text-sm text-gray-500 mt-1">Fill in vendor costs and the pricing engine handles markups, tax, fees, and P&L.</p>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left: Form */}
        <div className="lg:col-span-2 space-y-6">
          {/* Client Setup */}
          <div className="card">
            <h2 className="text-lg font-semibold text-gray-900 mb-4">Event Details</h2>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Estimate Name</label>
                <input
                  className="input-field"
                  value={form.estimate_name}
                  onChange={(e) => setForm({ ...form, estimate_name: e.target.value })}
                  placeholder="e.g., Pfizer CLT Dinner 2026"
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Client Name</label>
                <input
                  className="input-field"
                  value={form.client_name}
                  onChange={(e) => setForm({ ...form, client_name: e.target.value })}
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Event Name</label>
                <input
                  className="input-field"
                  value={form.event_name}
                  onChange={(e) => setForm({ ...form, event_name: e.target.value })}
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Event Date</label>
                <input
                  type="date"
                  className="input-field"
                  value={form.event_date}
                  onChange={(e) => setForm({ ...form, event_date: e.target.value })}
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Guest Count</label>
                <input
                  type="number"
                  className="input-field"
                  value={form.guest_count || ''}
                  onChange={(e) => setForm({ ...form, guest_count: parseInt(e.target.value) || 0 })}
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Location</label>
                <input
                  className="input-field"
                  value={form.location}
                  onChange={(e) => setForm({ ...form, location: e.target.value })}
                  placeholder="Charlotte, NC"
                />
              </div>
            </div>

            <div className="grid grid-cols-3 gap-4 mt-4 pt-4 border-t border-gray-100">
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Tax Rate</label>
                <input
                  type="number"
                  step="0.001"
                  className="input-field"
                  value={form.tax_rate}
                  onChange={(e) => setForm({ ...form, tax_rate: parseFloat(e.target.value) || 0 })}
                />
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Commission Scenario</label>
                <select
                  className="input-field"
                  value={form.commission_scenario}
                  onChange={(e) => setForm({ ...form, commission_scenario: e.target.value })}
                >
                  {COMMISSION_OPTIONS.map((opt) => (
                    <option key={opt.key} value={opt.key}>
                      {opt.label}
                    </option>
                  ))}
                </select>
              </div>
              <div>
                <label className="block text-xs font-medium text-gray-500 mb-1">Est. OpEx Hours</label>
                <input
                  type="number"
                  className="input-field"
                  value={form.opex_hours || ''}
                  onChange={(e) => setForm({ ...form, opex_hours: parseFloat(e.target.value) || 0 })}
                />
              </div>
            </div>
          </div>

          {/* Line Items by Section */}
          {Array.from(sections.entries()).map(([sectionName, sectionData]) => (
            <div key={sectionName} className="card">
              <div className="flex items-center justify-between mb-3">
                <div>
                  <h3 className="font-semibold text-gray-900">{sectionName}</h3>
                  <span className="text-xs text-gray-500">
                    {categoryLabel(sectionData.categoryKey)} &middot; {categoryMarkup(sectionData.categoryKey)}% markup
                    {!sectionData.isTaxable && ' \u00b7 Non-taxable'}
                  </span>
                </div>
                <button
                  onClick={() => addLineItem(sectionName, sectionData.categoryKey, sectionData.isTaxable)}
                  className="text-xs text-qc-600 hover:text-qc-800 font-medium"
                >
                  + Add Row
                </button>
              </div>

              <table className="w-full text-sm">
                <thead>
                  <tr className="text-xs text-gray-500 uppercase tracking-wider border-b border-gray-100">
                    <th className="text-left py-2 pr-2 w-1/3">Item</th>
                    <th className="text-right py-2 px-2 w-20">Qty</th>
                    <th className="text-right py-2 px-2 w-28">Unit Cost</th>
                    <th className="text-right py-2 px-2 w-28">Vendor Total</th>
                    <th className="text-right py-2 px-2 w-28">Client Price</th>
                    <th className="py-2 pl-2 w-8"></th>
                  </tr>
                </thead>
                <tbody>
                  {sectionData.items.map(({ item, globalIndex }) => {
                    const vendorCost =
                      item.item_type === 'flat_fee' ? item.unit_price : item.quantity * item.unit_price
                    const markup = categoryMarkup(item.category_key) / 100
                    const clientCost = vendorCost * (1 + markup)
                    return (
                      <tr key={globalIndex} className="border-b border-gray-50 hover:bg-gray-50">
                        <td className="py-2 pr-2">
                          <input
                            className="w-full border-0 bg-transparent text-sm focus:ring-0 p-0 outline-none"
                            value={item.name}
                            onChange={(e) => updateLineItem(globalIndex, 'name', e.target.value)}
                            placeholder="Item name"
                          />
                        </td>
                        <td className="py-2 px-2">
                          {item.item_type !== 'flat_fee' && (
                            <input
                              type="number"
                              className="w-full text-right border-0 bg-transparent text-sm focus:ring-0 p-0 outline-none"
                              value={item.quantity || ''}
                              onChange={(e) =>
                                updateLineItem(globalIndex, 'quantity', parseInt(e.target.value) || 0)
                              }
                            />
                          )}
                        </td>
                        <td className="py-2 px-2">
                          <input
                            type="number"
                            step="0.01"
                            className="w-full text-right border-0 bg-transparent text-sm focus:ring-0 p-0 outline-none"
                            value={item.unit_price || ''}
                            onChange={(e) =>
                              updateLineItem(globalIndex, 'unit_price', parseFloat(e.target.value) || 0)
                            }
                            placeholder="0.00"
                          />
                        </td>
                        <td className="py-2 px-2 text-right text-gray-600">{vendorCost > 0 ? fmt(vendorCost) : '-'}</td>
                        <td className="py-2 px-2 text-right font-medium text-gray-900">
                          {clientCost > 0 ? fmt(clientCost) : '-'}
                        </td>
                        <td className="py-2 pl-2 text-center">
                          <button
                            onClick={() => removeLineItem(globalIndex)}
                            className="text-gray-400 hover:text-red-500 text-xs"
                            title="Remove"
                          >
                            &times;
                          </button>
                        </td>
                      </tr>
                    )
                  })}
                </tbody>
              </table>
            </div>
          ))}

          {/* Action buttons */}
          <div className="flex gap-3">
            <button onClick={handleCalculate} className="btn-primary" disabled={saving}>
              {saving ? 'Calculating...' : 'Calculate & Save'}
            </button>
            {result?.id && (
              <button onClick={handleSaveAndView} className="btn-secondary">
                View Estimate
              </button>
            )}
          </div>
        </div>

        {/* Right: P&L and Summary */}
        <div className="space-y-6">
          {result ? (
            <>
              <PnLPanel pnl={result.pnl} />

              <div className="card">
                <h3 className="text-lg font-semibold text-gray-900 mb-3">Client-Ready Summary</h3>
                <div className="space-y-2">
                  {Object.entries(result.client_ready.sections).map(([section, amount]) => (
                    <div key={section} className="flex justify-between text-sm">
                      <span className="text-gray-600">{section}</span>
                      <span className="font-medium">{fmt(amount as number)}</span>
                    </div>
                  ))}
                  <div className="flex justify-between text-sm border-t border-gray-200 pt-2">
                    <span className="text-gray-600">Tax</span>
                    <span className="font-medium">{fmt(result.client_ready.tax)}</span>
                  </div>
                  <div className="flex justify-between text-base font-bold border-t-2 border-gray-900 pt-2">
                    <span>Total</span>
                    <span>{fmt(result.client_ready.total)}</span>
                  </div>
                </div>
              </div>
            </>
          ) : (
            <div className="card text-center py-12">
              <p className="text-sm text-gray-500">
                Enter vendor costs and click <strong>Calculate</strong> to see the P&L analysis.
              </p>
            </div>
          )}

          {/* Category Markup Reference */}
          <div className="card">
            <h3 className="text-sm font-semibold text-gray-700 mb-3">Markup Reference</h3>
            <div className="space-y-1.5">
              {categories.map((cat) => (
                <div key={cat.key} className="flex justify-between text-xs">
                  <span className="text-gray-600">{cat.label}</span>
                  <span className="font-mono font-medium text-gray-900">{cat.markup_pct}%</span>
                </div>
              ))}
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
