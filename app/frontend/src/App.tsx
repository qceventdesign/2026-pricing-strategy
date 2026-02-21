import { Routes, Route, Link, useLocation } from "react-router-dom";
import Dashboard from "./pages/Dashboard";
import NewEstimate from "./pages/NewEstimate";
import EstimateDetail from "./pages/EstimateDetail";
import ConfigView from "./pages/ConfigView";

const NAV = [
  { path: "/", label: "Dashboard" },
  { path: "/config", label: "Pricing Rules" },
];

export default function App() {
  const location = useLocation();

  return (
    <div className="min-h-screen">
      {/* Header */}
      <header className="border-b border-brand-100 bg-white">
        <div className="mx-auto flex max-w-6xl items-center justify-between px-6 py-4">
          <Link to="/" className="flex items-center gap-3">
            <div className="flex h-9 w-9 items-center justify-center rounded-lg bg-brand-700 text-sm font-bold text-white">
              QC
            </div>
            <div>
              <div className="font-semibold leading-tight text-brand-700">
                Quill Creative
              </div>
              <div className="text-xs text-brand-400">Pricing Engine</div>
            </div>
          </Link>
          <nav className="flex gap-1">
            {NAV.map((n) => (
              <Link
                key={n.path}
                to={n.path}
                className={`rounded-lg px-3 py-2 text-sm font-medium transition-colors ${
                  location.pathname === n.path
                    ? "bg-brand-100 text-brand-700"
                    : "text-brand-400 hover:text-brand-600"
                }`}
              >
                {n.label}
              </Link>
            ))}
          </nav>
        </div>
      </header>

      {/* Content */}
      <main className="mx-auto max-w-6xl px-6 py-8">
        <Routes>
          <Route path="/" element={<Dashboard />} />
          <Route path="/new/:type" element={<NewEstimate />} />
          <Route path="/estimate/:id" element={<EstimateDetail />} />
          <Route path="/config" element={<ConfigView />} />
        </Routes>
      </main>
    </div>
  );
}
