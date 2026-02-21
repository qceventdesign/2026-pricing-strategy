import type { HealthIndicator } from '../types'

const LEVEL_STYLES: Record<string, string> = {
  strong: 'health-strong',
  on_target: 'health-on-target',
  review: 'health-review',
  below_floor: 'health-below',
  thin: 'health-review',
  losing_money: 'health-below',
}

export default function HealthBadge({ health }: { health: HealthIndicator }) {
  const style = LEVEL_STYLES[health.level] || 'bg-gray-100 text-gray-700'
  return (
    <span
      className={`inline-flex items-center px-2.5 py-1 rounded-md text-xs font-semibold border ${style}`}
      title={health.meaning}
    >
      {health.label}
    </span>
  )
}
