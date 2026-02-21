import { useState, useEffect } from 'react'
import { useParams, Link, useNavigate } from 'react-router-dom'
import { api } from '../api/client'
import type { CalculationResult } from '../types'
import PnLPanel from '../components/PnLPanel'

function fmt(n: number): string {
  return n.toLocaleString('en-US', { style: 'currency', currency: 'USD' })
}

export default function EstimateDetail() {
  const { id } = useParams<{ id: string }>()
  const navigate = useNavigate()
  const [data, setData] = useState<CalculationResult | null>(null)
  const [audit, setAudit] = useState<any[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    if (!id) return
    const numId = parseInt(id)
    Promise.all([api.getEstimate(numId), api.getAuditLog(numId)])
      .then(([est, log]) => {
        setData(est)
        setAudit(log)
      })
      .finally(() => setLoading(false))
  }, [id])

  const handleDelete = async () => {
    if (!confirm('Delete this estimate? This cannot be undone.')) return
    await api.deleteEstimate(parseInt(id!))
    navigate('/')
  }

  if (loading) return <div className="card text-center py-12 text-gray-500">Loading...</div>
  if (!data) return <div className="card text-center py-12 text-red-500">Estimate not found</div>

  return (
    <div>
      <div className="flex items-center justify-between mb-6">
        <div>
          <Link to="/" className="text-sm text-qc-600 hover:text-qc-800 mb-1 inline-block">
            &larr; All Estimates
          </Link>
          <h1 className="text-2xl font-bold text-gray-900">{data.estimate_name}</h1>
          <p className="text-sm text-gray-500 mt-1">
            {data.client_name} &middot; {data.event_date || 'No date'} &middot;{' '}
            <span className="capitalize">{data.estimate_type}</span>
          </p>
        </div>
        <div className="flex gap-2">
          <a href={api.exportExcelUrl(parseInt(id!))} className="btn-secondary text-sm" download>
            Export Excel
          </a>
          <button onClick={handleDelete} className="text-sm text-red-600 hover:text-red-800 px-3 py-2">
            Delete
          </button>
        </div>
      </div>

      <div className="grid grid-cols-1 lg:grid-cols-3 gap-6">
        {/* Left: Details */}
        <div className="lg:col-span-2 space-y-6">
          {/* Event Info */}
          <div className="card">
            <h2 className="font-semibold text-gray-900 mb-3">Event Details</h2>
            <div className="grid grid-cols-3 gap-4 text-sm">
              <div>
                <span className="text-gray-500 block text-xs mb-0.5">Client</span>
                <span className="font-medium">{data.client_name || '-'}</span>
              </div>
              <div>
                <span className="text-gray-500 block text-xs mb-0.5">Event</span>
                <span className="font-medium">{data.event_name || '-'}</span>
              </div>
              <div>
                <span className="text-gray-500 block text-xs mb-0.5">Guests</span>
                <span className="font-medium">{data.guest_count || '-'}</span>
              </div>
              <div>
                <span className="text-gray-500 block text-xs mb-0.5">Location</span>
                <span className="font-medium">{data.location || '-'}</span>
              </div>
              <div>
                <span className="text-gray-500 block text-xs mb-0.5">Commission</span>
                <span className="font-medium">{data.commission_scenario.label}</span>
              </div>
              <div>
                <span className="text-gray-500 block text-xs mb-0.5">Status</span>
                <span className="font-medium capitalize">{data.status || 'draft'}</span>
              </div>
            </div>
          </div>

          {/* Line Items by Section */}
          {Object.entries(data.sections).map(([sectionName, sectionData]) => (
            <div key={sectionName} className="card">
              <div className="flex items-center justify-between mb-3">
                <h3 className="font-semibold text-gray-900">{sectionName}</h3>
                <span className="text-sm font-medium text-gray-700">{fmt(sectionData.client_cost)}</span>
              </div>
              <table className="w-full text-sm">
                <thead>
                  <tr className="text-xs text-gray-500 uppercase tracking-wider border-b border-gray-100">
                    <th className="text-left py-2">Item</th>
                    <th className="text-right py-2">Qty</th>
                    <th className="text-right py-2">Vendor Cost</th>
                    <th className="text-right py-2">Markup</th>
                    <th className="text-right py-2">Client Price</th>
                  </tr>
                </thead>
                <tbody>
                  {(sectionData as any).items.map((item: any, i: number) => (
                    <tr key={i} className="border-b border-gray-50">
                      <td className="py-2">
                        <span className="font-medium">{item.name}</span>
                        <span className="text-xs text-gray-400 ml-2">{item.category_label}</span>
                      </td>
                      <td className="py-2 text-right text-gray-600">
                        {item.item_type !== 'flat_fee' ? item.quantity : '-'}
                      </td>
                      <td className="py-2 text-right text-gray-600">{fmt(item.vendor_cost)}</td>
                      <td className="py-2 text-right text-gray-500">{(item.markup_pct * 100).toFixed(0)}%</td>
                      <td className="py-2 text-right font-medium">{fmt(item.client_cost)}</td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          ))}

          {/* Subtotals */}
          <div className="card">
            <h3 className="font-semibold text-gray-900 mb-3">Invoice Summary</h3>
            <div className="space-y-2 text-sm">
              <div className="flex justify-between">
                <span className="text-gray-600">Taxable Subtotal (Client)</span>
                <span>{fmt(data.subtotals.taxable_client)}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-gray-600">Non-Taxable Subtotal (Client)</span>
                <span>{fmt(data.subtotals.nontaxable_client)}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-gray-600">Sales Tax</span>
                <span>{fmt(data.subtotals.sales_tax)}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-gray-600">CC Processing Fee</span>
                <span>{fmt(data.subtotals.cc_fee)}</span>
              </div>
              <div className="flex justify-between font-bold border-t-2 border-gray-900 pt-2 text-base">
                <span>Client Invoice Total</span>
                <span>{fmt(data.subtotals.client_invoice_total)}</span>
              </div>
            </div>
          </div>

          {/* Audit Log */}
          {audit.length > 0 && (
            <div className="card">
              <h3 className="font-semibold text-gray-900 mb-3">Audit Log</h3>
              <div className="space-y-2">
                {audit.map((log) => (
                  <div key={log.id} className="flex items-center justify-between text-sm py-1 border-b border-gray-50">
                    <span className="capitalize text-gray-700">{log.action}</span>
                    <span className="text-xs text-gray-400">
                      {new Date(log.created_at).toLocaleString()}
                    </span>
                  </div>
                ))}
              </div>
            </div>
          )}
        </div>

        {/* Right: P&L + Client Ready */}
        <div className="space-y-6">
          <PnLPanel pnl={data.pnl} />

          <div className="card">
            <h3 className="text-lg font-semibold text-gray-900 mb-3">Client-Ready Summary</h3>
            <p className="text-xs text-gray-500 mb-3">Rounded to nearest $50 for presentation</p>
            <div className="space-y-2">
              {Object.entries(data.client_ready.sections).map(([section, amount]) => (
                <div key={section} className="flex justify-between text-sm">
                  <span className="text-gray-600">{section}</span>
                  <span className="font-medium">{fmt(amount as number)}</span>
                </div>
              ))}
              <div className="flex justify-between text-sm border-t border-gray-200 pt-2">
                <span className="text-gray-600">Tax</span>
                <span className="font-medium">{fmt(data.client_ready.tax)}</span>
              </div>
              <div className="flex justify-between text-base font-bold border-t-2 border-gray-900 pt-2">
                <span>Total</span>
                <span>{fmt(data.client_ready.total)}</span>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}
