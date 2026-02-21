export interface LineItem {
  description: string;
  category_key: string;
  vendor_cost: number;
  quantity: number;
  is_taxable: boolean;
}

export interface CalculatedLineItem extends LineItem {
  category_label: string;
  vendor_total: number;
  markup_pct: number;
  markup_amount: number;
  client_cost: number;
  tax_rate_pct: number;
  tax_amount: number;
  client_total: number;
}

export interface Summary {
  total_vendor_cost: number;
  total_client_cost: number;
  management_fee: number;
  client_invoice_subtotal: number;
  total_tax: number;
  client_invoice_total: number;
  commission_scenario: string;
  commission_label: string;
  commission_pct: number;
  commission_amount: number;
  cc_fee_pct: number;
  cc_fee_amount: number;
  qc_gross_revenue: number;
  qc_gross_margin_pct: number;
  opex_hours: number;
  opex_rate: number;
  opex_cost: number;
  travel_expenses: number;
  true_net_profit: number;
  true_net_margin_pct: number;
}

export interface HealthIndicator {
  label: string;
  meaning: string;
  margin_pct: number;
}

export interface Health {
  qc_gross_margin: HealthIndicator;
  true_net_margin: HealthIndicator;
}

export interface Estimate {
  id: string;
  created_at: string;
  updated_at: string;
  name: string;
  estimate_type: string;
  client_name: string;
  event_name: string;
  event_date: string;
  guest_count: number;
  commission_scenario: string;
  opex_hours: number;
  management_fee: number;
  travel_expenses: number;
  tax_rate_pct: number;
  line_items: LineItem[];
  notes: string;
  calculated?: Summary | { line_items: CalculatedLineItem[]; summary: Summary; health: Health };
  health?: Health;
}

export interface CategoryOption {
  key: string;
  label: string;
  markup_pct: number;
  rationale: string;
}

export interface EstimateTemplate {
  label: string;
  default_categories: string[];
}

export interface Config {
  markups: any;
  rates: any;
  fees: any;
  commissions: any;
  thresholds: any;
}
