export interface LineItem {
  id?: number
  section: string
  name: string
  category_key: string
  item_type: 'quantity' | 'flat_fee' | 'per_person'
  unit_price: number
  quantity: number
  is_taxable: boolean
  notes: string
  sort_order: number
  // Calculated (returned from API)
  vendor_cost?: number
  markup_pct?: number
  client_cost?: number
  markup_amount?: number
  category_label?: string
}

export interface EstimateInput {
  estimate_name: string
  estimate_type: string
  client_name: string
  event_name: string
  event_date: string
  event_time: string
  guest_count: number
  location: string
  tax_rate: number
  commission_scenario: string
  cc_processing_pct: number
  gratuity_pct: number
  admin_fee_pct: number
  gdp_enabled: boolean
  opex_hours: number
  line_items: LineItem[]
}

export interface HealthIndicator {
  level: string
  label: string
  meaning: string
}

export interface PnL {
  client_invoice_total: number
  commission_amount: number
  gdp_fee: number
  total_vendor_costs: number
  qc_gross_revenue: number
  qc_margin_pct: number
  qc_margin_health: HealthIndicator
  opex_hours: number
  opex_rate: number
  opex_cost: number
  true_net_profit: number
  true_net_margin_pct: number
  true_net_health: HealthIndicator
}

export interface CalculationResult {
  id?: number
  estimate_name: string
  estimate_type: string
  client_name: string
  event_name: string
  event_date: string
  event_time: string
  guest_count: number
  location: string
  status?: string
  commission_scenario: {
    key: string
    label: string
    commission_pct: number
  }
  line_items: LineItem[]
  sections: Record<string, { vendor_cost: number; client_cost: number; items: LineItem[] }>
  subtotals: {
    taxable_vendor: number
    taxable_client: number
    nontaxable_vendor: number
    nontaxable_client: number
    total_vendor: number
    total_client_pretax: number
    sales_tax: number
    cc_fee: number
    commission_amount: number
    client_invoice_total: number
  }
  pnl: PnL
  client_ready: {
    sections: Record<string, number>
    tax: number
    total: number
  }
}

export interface EstimateSummary {
  id: number
  estimate_name: string
  estimate_type: string
  client_name: string
  event_name: string
  event_date: string
  status: string
  created_at: string
  updated_at: string
}

export interface Category {
  key: string
  label: string
  markup_pct: number
  rationale: string
}

export interface EstimateType {
  key: string
  label: string
}

export interface TemplateSection {
  name: string
  category_key: string
  is_taxable: boolean
  default_items: { name: string; item_type: string }[]
}

export interface EstimateTemplate {
  label: string
  sections: TemplateSection[]
}

export interface CommissionScenario {
  label: string
  commission_pct: number
  cc_fee_pct: number
  net_to_qc_pct: number
}
