const BASE = "/api";

async function request<T>(path: string, options?: RequestInit): Promise<T> {
  const res = await fetch(`${BASE}${path}`, {
    headers: { "Content-Type": "application/json" },
    ...options,
  });
  if (!res.ok) {
    const err = await res.json().catch(() => ({ detail: res.statusText }));
    throw new Error(err.detail || "Request failed");
  }
  // For delete, may not have JSON body
  if (res.status === 204) return {} as T;
  return res.json();
}

export const api = {
  // Config
  getConfig: () => request<any>("/config"),
  getCategories: () => request<any[]>("/config/categories"),
  getTemplates: () => request<Record<string, any>>("/config/templates"),

  // Calculate (stateless)
  calculate: (data: any) => request<any>("/calculate", { method: "POST", body: JSON.stringify(data) }),

  // Estimates CRUD
  listEstimates: () => request<any[]>("/estimates"),
  getEstimate: (id: string) => request<any>(`/estimates/${id}`),
  createEstimate: (data: any) => request<any>("/estimates", { method: "POST", body: JSON.stringify(data) }),
  updateEstimate: (id: string, data: any) =>
    request<any>(`/estimates/${id}`, { method: "PUT", body: JSON.stringify(data) }),
  deleteEstimate: (id: string) => request<any>(`/estimates/${id}`, { method: "DELETE" }),

  // Export
  exportExcelUrl: (id: string) => `${BASE}/estimates/${id}/export/excel`,
};
