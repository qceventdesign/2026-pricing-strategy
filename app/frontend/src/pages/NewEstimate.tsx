import { useEffect, useState, useCallback } from "react";
import { useParams, useNavigate } from "react-router-dom";
import { api } from "../api";
import type { CategoryOption } from "../types";
import LineItemRow from "../components/LineItemRow";
import PnLPanel from "../components/PnLPanel";

const TYPE_LABELS: Record<string, string> = {
  venue: "Venue Estimate",
  decor: "Decor Estimate",
  transportation: "Transportation Estimate",
  entertainment: "Entertainment Estimate",
  tour: "Tour Estimate",
};

interface FormLineItem {
  description: string;
  category_key: string;
  vendor_cost: number;
  quantity: number;
  is_taxable: boolean;
}

export default function NewEstimate() {
  const { type } = useParams<{ type: string }>();
  const navigate = useNavigate();

  const [categories, setCategories] = useState<CategoryOption[]>([]);
  const [templates, setTemplates] = useState<Record<string, any>>({});
  const [saving, setSaving] = useState(false);

  // Form state
  const [name, setName] = useState("");
  const [clientName, setClientName] = useState("");
  const [eventName, setEventName] = useState("");
  const [eventDate, setEventDate] = useState("");
  const [guestCount, setGuestCount] = useState(0);
  const [commission, setCommission] = useState("direct_client");
  const [opexHours, setOpexHours] = useState(0);
  const [managementFee, setManagementFee] = useState(0);
  const [travelExpenses, setTravelExpenses] = useState(0);
  const [taxRate, setTaxRate] = useState(0);
  const [notes, setNotes] = useState("");
  const [lineItems, setLineItems] = useState<FormLineItem[]>([]);

  // Live calculation
  const [calcResult, setCalcResult] = useState<any>(null);

  useEffect(() => {
    Promise.all([api.getCategories(), api.getTemplates()]).then(
      ([cats, tmpl]) => {
        setCategories(cats);
        setTemplates(tmpl);

        // Pre-populate line items from template
        if (type && tmpl[type]) {
          const defaultCats = tmpl[type].default_categories as string[];
          setLineItems(
            defaultCats.map((ck) => ({
              description: "",
              category_key: ck,
              vendor_cost: 0,
              quantity: 1,
              is_taxable: false,
            }))
          );
          setName(TYPE_LABELS[type] || "");
        }
      }
    );
  }, [type]);

  // Live calculate on changes
  const recalc = useCallback(async () => {
    const validItems = lineItems.filter(
      (li) => li.category_key && li.vendor_cost > 0
    );
    if (validItems.length === 0) {
      setCalcResult(null);
      return;
    }
    try {
      const result = await api.calculate({
        line_items: validItems,
        commission_scenario: commission,
        opex_hours: opexHours,
        management_fee: managementFee,
        travel_expenses: travelExpenses,
        tax_rate_pct: taxRate,
      });
      setCalcResult(result);
    } catch {
      // Ignore errors during live calc
    }
  }, [lineItems, commission, opexHours, managementFee, travelExpenses, taxRate]);

  useEffect(() => {
    const timer = setTimeout(recalc, 300);
    return () => clearTimeout(timer);
  }, [recalc]);

  const handleLineItemChange = (index: number, field: string, value: any) => {
    setLineItems((prev) => {
      const next = [...prev];
      next[index] = { ...next[index], [field]: value };
      return next;
    });
  };

  const addLineItem = () => {
    setLineItems((prev) => [
      ...prev,
      { description: "", category_key: "", vendor_cost: 0, quantity: 1, is_taxable: false },
    ]);
  };

  const removeLineItem = (index: number) => {
    setLineItems((prev) => prev.filter((_, i) => i !== index));
  };

  const handleSave = async () => {
    setSaving(true);
    try {
      const est = await api.createEstimate({
        name,
        estimate_type: type || "venue",
        client_name: clientName,
        event_name: eventName,
        event_date: eventDate,
        guest_count: guestCount,
        commission_scenario: commission,
        opex_hours: opexHours,
        management_fee: managementFee,
        travel_expenses: travelExpenses,
        tax_rate_pct: taxRate,
        line_items: lineItems.filter((li) => li.category_key && li.vendor_cost > 0),
        notes,
      });
      navigate(`/estimate/${est.id}`);
    } catch (err: any) {
      alert(err.message || "Failed to save");
    } finally {
      setSaving(false);
    }
  };

  return (
    <div className="space-y-6">
      <div>
        <h1 className="text-3xl font-bold text-brand-700">
          {TYPE_LABELS[type || ""] || "New Estimate"}
        </h1>
        <p className="mt-1 text-brand-400">
          Add line items, set vendor costs, and see your margins live.
        </p>
      </div>

      <div className="grid grid-cols-3 gap-6">
        {/* Left: form (2 cols) */}
        <div className="col-span-2 space-y-6">
          {/* Event info */}
          <div className="card">
            <h2 className="mb-4 text-lg font-semibold text-brand-700">
              Event Details
            </h2>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="label">Estimate Name</label>
                <input
                  className="input"
                  value={name}
                  onChange={(e) => setName(e.target.value)}
                />
              </div>
              <div>
                <label className="label">Client Name</label>
                <input
                  className="input"
                  value={clientName}
                  onChange={(e) => setClientName(e.target.value)}
                />
              </div>
              <div>
                <label className="label">Event Name</label>
                <input
                  className="input"
                  value={eventName}
                  onChange={(e) => setEventName(e.target.value)}
                />
              </div>
              <div>
                <label className="label">Event Date</label>
                <input
                  type="date"
                  className="input"
                  value={eventDate}
                  onChange={(e) => setEventDate(e.target.value)}
                />
              </div>
              <div>
                <label className="label">Guest Count</label>
                <input
                  type="number"
                  className="input"
                  min={0}
                  value={guestCount || ""}
                  onChange={(e) => setGuestCount(parseInt(e.target.value) || 0)}
                />
              </div>
              <div>
                <label className="label">Commission Scenario</label>
                <select
                  className="input"
                  value={commission}
                  onChange={(e) => setCommission(e.target.value)}
                >
                  <option value="direct_client">Direct Client (0%)</option>
                  <option value="gdp_standard">GDP Standard (6.5%)</option>
                  <option value="gdp_plus_client">
                    GDP + Client (21.5%)
                  </option>
                </select>
              </div>
            </div>
          </div>

          {/* Line items */}
          <div className="card">
            <div className="mb-4 flex items-center justify-between">
              <h2 className="text-lg font-semibold text-brand-700">
                Line Items
              </h2>
              <button onClick={addLineItem} className="btn-secondary text-sm">
                + Add Item
              </button>
            </div>
            <div className="overflow-x-auto">
              <table className="w-full text-sm">
                <thead>
                  <tr className="border-b border-brand-100">
                    <th className="pb-2 text-left font-medium text-brand-500">
                      Description
                    </th>
                    <th className="pb-2 text-left font-medium text-brand-500">
                      Category
                    </th>
                    <th className="pb-2 text-center font-medium text-brand-500">
                      Qty
                    </th>
                    <th className="pb-2 text-right font-medium text-brand-500">
                      Vendor $
                    </th>
                    <th className="pb-2 text-center font-medium text-brand-500">
                      Markup
                    </th>
                    <th className="pb-2 text-right font-medium text-brand-500">
                      Client $
                    </th>
                    <th className="pb-2 text-center font-medium text-brand-500">
                      Tax
                    </th>
                    <th className="pb-2"></th>
                  </tr>
                </thead>
                <tbody>
                  {lineItems.map((item, i) => (
                    <LineItemRow
                      key={i}
                      index={i}
                      item={item}
                      categories={categories}
                      onChange={handleLineItemChange}
                      onRemove={removeLineItem}
                    />
                  ))}
                </tbody>
              </table>
              {lineItems.length === 0 && (
                <p className="py-6 text-center text-sm text-brand-300">
                  No line items yet. Click "+ Add Item" to get started.
                </p>
              )}
            </div>
          </div>

          {/* Additional costs */}
          <div className="card">
            <h2 className="mb-4 text-lg font-semibold text-brand-700">
              Additional Costs & Fees
            </h2>
            <div className="grid grid-cols-2 gap-4">
              <div>
                <label className="label">OpEx Hours</label>
                <input
                  type="number"
                  className="input"
                  min={0}
                  value={opexHours || ""}
                  onChange={(e) => setOpexHours(parseFloat(e.target.value) || 0)}
                />
                <p className="mt-1 text-xs text-brand-300">
                  Internal team time at $90/hr
                </p>
              </div>
              <div>
                <label className="label">Management Fee ($)</label>
                <input
                  type="number"
                  className="input"
                  min={0}
                  step={100}
                  value={managementFee || ""}
                  onChange={(e) =>
                    setManagementFee(parseFloat(e.target.value) || 0)
                  }
                />
              </div>
              <div>
                <label className="label">Travel Expenses ($)</label>
                <input
                  type="number"
                  className="input"
                  min={0}
                  value={travelExpenses || ""}
                  onChange={(e) =>
                    setTravelExpenses(parseFloat(e.target.value) || 0)
                  }
                />
              </div>
              <div>
                <label className="label">Tax Rate (%)</label>
                <input
                  type="number"
                  className="input"
                  min={0}
                  step={0.1}
                  value={taxRate || ""}
                  onChange={(e) => setTaxRate(parseFloat(e.target.value) || 0)}
                />
              </div>
            </div>
          </div>

          {/* Notes */}
          <div className="card">
            <label className="label">Notes</label>
            <textarea
              className="input"
              rows={3}
              value={notes}
              onChange={(e) => setNotes(e.target.value)}
              placeholder="Internal notes, special instructions..."
            />
          </div>

          {/* Actions */}
          <div className="flex gap-3">
            <button
              onClick={handleSave}
              disabled={saving}
              className="btn-primary"
            >
              {saving ? "Saving..." : "Save Estimate"}
            </button>
            <button onClick={() => navigate("/")} className="btn-secondary">
              Cancel
            </button>
          </div>
        </div>

        {/* Right: live P&L (1 col) */}
        <div className="col-span-1">
          <div className="sticky top-8">
            <h2 className="mb-4 text-lg font-semibold text-brand-700">
              Live P&L
            </h2>
            {calcResult ? (
              <PnLPanel
                summary={calcResult.summary}
                health={calcResult.health}
              />
            ) : (
              <div className="card py-8 text-center text-sm text-brand-300">
                Add line items to see live P&L
              </div>
            )}
          </div>
        </div>
      </div>
    </div>
  );
}
