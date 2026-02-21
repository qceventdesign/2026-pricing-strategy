import { Routes, Route, Link, useLocation } from 'react-router-dom'
import Dashboard from './pages/Dashboard'
import NewEstimate from './pages/NewEstimate'
import EstimateDetail from './pages/EstimateDetail'
import ConfigView from './pages/ConfigView'

const NAV_ITEMS = [
  { path: '/', label: 'Estimates' },
  { path: '/config', label: 'Pricing Config' },
]

export default function App() {
  const location = useLocation()

  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-white border-b border-gray-200 sticky top-0 z-10">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex items-center justify-between h-16">
            <Link to="/" className="flex items-center gap-3">
              <div className="w-8 h-8 bg-qc-600 rounded-lg flex items-center justify-center">
                <span className="text-white font-bold text-sm">QC</span>
              </div>
              <span className="font-semibold text-lg text-qc-800">Pricing Tool</span>
            </Link>
            <nav className="flex items-center gap-1">
              {NAV_ITEMS.map((item) => (
                <Link
                  key={item.path}
                  to={item.path}
                  className={`px-3 py-2 rounded-lg text-sm font-medium transition-colors ${
                    location.pathname === item.path
                      ? 'bg-qc-100 text-qc-800'
                      : 'text-gray-600 hover:text-qc-700 hover:bg-gray-100'
                  }`}
                >
                  {item.label}
                </Link>
              ))}
            </nav>
          </div>
        </div>
      </header>

      {/* Main */}
      <main className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <Routes>
          <Route path="/" element={<Dashboard />} />
          <Route path="/new/:type" element={<NewEstimate />} />
          <Route path="/estimate/:id" element={<EstimateDetail />} />
          <Route path="/config" element={<ConfigView />} />
        </Routes>
      </main>
    </div>
  )
}
