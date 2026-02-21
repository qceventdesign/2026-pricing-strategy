const BASE = '/api'

async function request<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    headers: { 'Content-Type': 'application/json' },
    ...options,
  })
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }))
    throw new Error(err.detail || 'Request failed')
  }
  return res.json()
}

export const api = {
  // Config
  getConfig: () => request<any>('/config/'),
  getCategories: () => request<any[]>('/config/markups/categories'),
  getCommissions: () => request<any>('/config/commissions'),

  // Estimate types & templates
  getEstimateTypes: () => request<any[]>('/estimates/types'),
  getTemplate: (type: string) => request<any>(`/estimates/types/${type}/template`),

  // Estimates CRUD
  listEstimates: () => request<any[]>('/estimates/'),
  getEstimate: (id: number) => request<any>(`/estimates/${id}`),
  createEstimate: (data: any) => request<any>('/estimates/', { method: 'POST', body: JSON.stringify(data) }),
  updateEstimate: (id: number, data: any) => request<any>(`/estimates/${id}`, { method: 'PUT', body: JSON.stringify(data) }),
  deleteEstimate: (id: number) => request<void>(`/estimates/${id}`, { method: 'DELETE' }),
  getAuditLog: (id: number) => request<any[]>(`/estimates/${id}/audit`),

  // Export
  exportExcelUrl: (id: number) => `${BASE}/export/${id}/excel`,
}
