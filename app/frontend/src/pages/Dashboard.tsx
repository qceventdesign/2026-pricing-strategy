import { useState, useEffect } from 'react'
import { Link } from 'react-router-dom'
import { api } from '../api/client'
import type { EstimateSummary, EstimateType } from '../types'

const STATUS_STYLES: Record<string, string> = {
  draft: 'bg-gray-100 text-gray-700',
  sent: 'bg-blue-100 text-blue-700',
  approved: 'bg-emerald-100 text-emerald-700',
  archived: 'bg-gray-100 text-gray-500',
}

export default function Dashboard() {
  const [estimates, setEstimates] = useState<EstimateSummary[]>([])
  const [types, setTypes] = useState<EstimateType[]>([])
  const [loading, setLoading] = useState(true)

  useEffect(() => {
    Promise.all([api.listEstimates(), api.getEstimateTypes()])
      .then(([est, t]) => {
        setEstimates(est)
        setTypes(t)
      })
      .finally(() => setLoading(false))
  }, [])

  return (
    <div>
      <div className="flex items-center justify-between mb-8">
        <div>
          <h1 className="text-2xl font-bold text-gray-900">Estimates</h1>
          <p className="text-sm text-gray-500 mt-1">Create and manage event pricing estimates</p>
        </div>
        <div className="relative group">
          <button className="btn-primary">+ New Estimate</button>
          <div className="absolute right-0 mt-2 w-56 bg-white rounded-lg shadow-lg border border-gray-200 py-1 hidden group-hover:block z-20">
            {types.map((t) => (
              <Link
                key={t.key}
                to={`/new/${t.key}`}
                className="block px-4 py-2 text-sm text-gray-700 hover:bg-qc-50 hover:text-qc-800"
              >
                {t.label}
              </Link>
            ))}
          </div>
        </div>
      </div>

      {loading ? (
        <div className="card text-center py-12 text-gray-500">Loading...</div>
      ) : estimates.length === 0 ? (
        <div className="card text-center py-16">
          <div className="text-gray-400 text-4xl mb-4">&#9634;</div>
          <h2 className="text-lg font-semibold text-gray-700 mb-2">No estimates yet</h2>
          <p className="text-sm text-gray-500 mb-6">
            Create your first estimate to start using the pricing tool.
          </p>
          <div className="flex flex-wrap gap-3 justify-center">
            {types.map((t) => (
              <Link key={t.key} to={`/new/${t.key}`} className="btn-secondary text-sm">
                {t.label}
              </Link>
            ))}
          </div>
        </div>
      ) : (
        <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
          <table className="w-full">
            <thead>
              <tr className="bg-gray-50 border-b border-gray-200">
                <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                  Estimate
                </th>
                <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                  Client
                </th>
                <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                  Type
                </th>
                <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                  Date
                </th>
                <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                  Status
                </th>
                <th className="text-left px-4 py-3 text-xs font-semibold text-gray-500 uppercase tracking-wider">
                  Updated
                </th>
              </tr>
            </thead>
            <tbody className="divide-y divide-gray-100">
              {estimates.map((est) => (
                <tr key={est.id} className="hover:bg-gray-50 transition-colors">
                  <td className="px-4 py-3">
                    <Link to={`/estimate/${est.id}`} className="font-medium text-qc-700 hover:text-qc-900">
                      {est.estimate_name}
                    </Link>
                  </td>
                  <td className="px-4 py-3 text-sm text-gray-600">{est.client_name || '-'}</td>
                  <td className="px-4 py-3 text-sm text-gray-600 capitalize">{est.estimate_type}</td>
                  <td className="px-4 py-3 text-sm text-gray-600">{est.event_date || '-'}</td>
                  <td className="px-4 py-3">
                    <span
                      className={`text-xs font-medium px-2 py-1 rounded-full capitalize ${
                        STATUS_STYLES[est.status] || STATUS_STYLES.draft
                      }`}
                    >
                      {est.status}
                    </span>
                  </td>
                  <td className="px-4 py-3 text-sm text-gray-500">
                    {est.updated_at ? new Date(est.updated_at).toLocaleDateString() : '-'}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </div>
  )
}
