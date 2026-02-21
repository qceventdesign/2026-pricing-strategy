import type { CategoryOption } from "../types";

interface Props {
  index: number;
  item: {
    description: string;
    category_key: string;
    vendor_cost: number;
    quantity: number;
    is_taxable: boolean;
  };
  categories: CategoryOption[];
  onChange: (index: number, field: string, value: any) => void;
  onRemove: (index: number) => void;
}

export default function LineItemRow({
  index,
  item,
  categories,
  onChange,
  onRemove,
}: Props) {
  const cat = categories.find((c) => c.key === item.category_key);
  const markupPct = cat?.markup_pct ?? 0;
  const vendorTotal = item.vendor_cost * item.quantity;
  const clientCost = vendorTotal * (1 + markupPct / 100);

  return (
    <tr className="border-b border-brand-50 hover:bg-brand-50/50">
      <td className="py-2 pr-2">
        <input
          type="text"
          className="input"
          placeholder="Description"
          value={item.description}
          onChange={(e) => onChange(index, "description", e.target.value)}
        />
      </td>
      <td className="py-2 pr-2">
        <select
          className="input"
          value={item.category_key}
          onChange={(e) => onChange(index, "category_key", e.target.value)}
        >
          <option value="">Select...</option>
          {categories.map((c) => (
            <option key={c.key} value={c.key}>
              {c.label} ({c.markup_pct}%)
            </option>
          ))}
        </select>
      </td>
      <td className="py-2 pr-2">
        <input
          type="number"
          className="input w-16 text-right"
          min={1}
          value={item.quantity}
          onChange={(e) => onChange(index, "quantity", parseInt(e.target.value) || 1)}
        />
      </td>
      <td className="py-2 pr-2">
        <input
          type="number"
          className="input w-28 text-right"
          min={0}
          step={0.01}
          value={item.vendor_cost || ""}
          placeholder="0.00"
          onChange={(e) =>
            onChange(index, "vendor_cost", parseFloat(e.target.value) || 0)
          }
        />
      </td>
      <td className="py-2 pr-2 text-center text-sm text-brand-400">
        {markupPct > 0 ? `${markupPct}%` : "-"}
      </td>
      <td className="py-2 pr-2 text-right text-sm font-medium">
        ${clientCost.toLocaleString("en-US", { minimumFractionDigits: 2 })}
      </td>
      <td className="py-2 pr-2 text-center">
        <input
          type="checkbox"
          checked={item.is_taxable}
          onChange={(e) => onChange(index, "is_taxable", e.target.checked)}
          className="h-4 w-4 rounded border-brand-300"
        />
      </td>
      <td className="py-2">
        <button
          onClick={() => onRemove(index)}
          className="text-red-400 hover:text-red-600 text-sm"
          title="Remove"
        >
          Remove
        </button>
      </td>
    </tr>
  );
}
