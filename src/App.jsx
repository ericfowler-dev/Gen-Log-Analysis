import { useEffect, useMemo, useRef, useState } from 'react'
import {
  Activity,
  AlertTriangle,
  BarChart3,
  CheckCircle2,
  FileSpreadsheet,
  Gauge,
  Loader2,
  Table2,
  Upload,
  Zap,
} from 'lucide-react'
import {
  CartesianGrid,
  Legend,
  Line,
  LineChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import { durationLabel, formatCompact, formatDateTime, formatNumber } from './lib/format.js'

const tabs = [
  { id: 'overview', label: 'Overview', icon: Gauge },
  { id: 'trends', label: 'Trends', icon: BarChart3 },
  { id: 'events', label: 'Events', icon: AlertTriangle },
  { id: 'parameters', label: 'Parameters', icon: Table2 },
  { id: 'raw', label: 'Raw Tabs', icon: Table2 },
]

const chartGroups = [
  {
    title: 'Engine',
    lines: [
      { key: 'rpm', name: 'RPM', color: '#2563eb' },
      { key: 'frequency', name: 'Frequency', color: '#16a34a' },
      { key: 'oilPressure', name: 'Oil Pressure', color: '#dc2626' },
    ],
  },
  {
    title: 'Temperatures',
    lines: [
      { key: 'coolantTemp', name: 'Coolant', color: '#ea580c' },
      { key: 'oilTemp', name: 'Oil', color: '#9333ea' },
      { key: 'intakeTemp', name: 'Intake', color: '#0891b2' },
    ],
  },
  {
    title: 'Electrical',
    lines: [
      { key: 'voltageAvg', name: 'Avg Voltage', color: '#0f766e' },
      { key: 'kw', name: 'kW', color: '#7c3aed' },
      { key: 'voltageImbalance', name: 'Voltage Imbalance %', color: '#be123c' },
    ],
  },
  {
    title: 'DiagSmart Engine',
    lines: [
      { key: 'calculatedLoad', name: 'Load %', color: '#b45309' },
      { key: 'batteryVoltage', name: 'Battery V', color: '#0f766e' },
      { key: 'railPressure', name: 'Rail Pressure', color: '#4f46e5' },
    ],
  },
  {
    title: 'Air / Exhaust',
    lines: [
      { key: 'turboSpeed', name: 'Turbo Speed', color: '#0369a1' },
      { key: 'turboTemp', name: 'Turbo Temp', color: '#c2410c' },
      { key: 'exhaustTempAvg', name: 'Avg Exhaust Temp', color: '#be123c' },
      { key: 'exhaustTempSpread', name: 'Exhaust Spread', color: '#64748b' },
    ],
  },
]

function App() {
  const [analysis, setAnalysis] = useState(null)
  const [activeTab, setActiveTab] = useState('overview')
  const [status, setStatus] = useState({ state: 'idle', message: '' })
  const [selectedRawSheet, setSelectedRawSheet] = useState(null)
  const fileInput = useRef(null)
  const workerRef = useRef(null)

  const eventCounts = useMemo(() => {
    const counts = { critical: 0, warning: 0, info: 0, state: 0, data: 0 }
    analysis?.events?.forEach((event) => {
      counts[event.severity] = (counts[event.severity] || 0) + 1
    })
    return counts
  }, [analysis])
  const visibleTabs = analysis?.error ? tabs.filter((tab) => tab.id === 'raw') : tabs

  async function handleFile(file) {
    if (!file) return
    setStatus({ state: 'loading', message: `Reading ${file.name}` })
    setAnalysis(null)

    const extension = (file.name.split('.').pop() || '').toLowerCase()
    const isWorkbook = ['xlsx', 'xls', 'xlsm'].includes(extension)
    const payload = {
      fileName: file.name,
      buffer: isWorkbook ? await file.arrayBuffer() : null,
      text: isWorkbook ? null : await file.text(),
    }

    workerRef.current?.terminate()
    const worker = new Worker(new URL('./workers/logWorker.js', import.meta.url), { type: 'module' })
    workerRef.current = worker
    worker.onmessage = (event) => {
      const { type, result, message } = event.data
      worker.terminate()
      if (type === 'error') {
        setStatus({ state: 'error', message })
        return
      }
      setAnalysis(result)
      setSelectedRawSheet(result.sheets?.[0]?.name || null)
      setActiveTab(result.error ? 'raw' : 'overview')
      setStatus({
        state: result.error ? 'error' : 'ready',
        message: result.error || `Analyzed ${result.stats.sampleCount.toLocaleString()} timestamped samples from ${result.selectedSheet}.`,
      })
    }
    worker.postMessage(payload, payload.buffer ? [payload.buffer] : [])
  }

  function handleDrop(event) {
    event.preventDefault()
    handleFile(event.dataTransfer.files?.[0])
  }

  return (
    <div className="app-shell">
      <header className="topbar">
        <div>
          <p className="eyebrow">WinScope / Gen-Set</p>
          <h1>Gen-Set Log Analyzer</h1>
        </div>
        <button className="primary-action" onClick={() => fileInput.current?.click()}>
          <Upload size={18} />
          Upload Log
        </button>
        <input
          ref={fileInput}
          className="hidden-input"
          type="file"
          accept=".xlsx,.xls,.xlsm,.csv,.txt,.log"
          onChange={(event) => handleFile(event.target.files?.[0])}
        />
      </header>

      <main>
        <section className="upload-band" onDragOver={(event) => event.preventDefault()} onDrop={handleDrop}>
          <div className="upload-copy">
            <FileSpreadsheet size={34} />
            <div>
              <h2>Upload an unformatted WinScope export or Gen-Set log</h2>
              <p>Excel workbooks, CSV files, tab-delimited logs, and plain text tables are parsed locally in your browser.</p>
            </div>
          </div>
          <div className={`status-pill ${status.state}`}>
            {status.state === 'loading' ? <Loader2 size={16} className="spin" /> : status.state === 'ready' ? <CheckCircle2 size={16} /> : <Activity size={16} />}
            <span>{status.message || 'Waiting for a file'}</span>
          </div>
        </section>

        {analysis ? (
          <>
            <nav className="tabbar" aria-label="Analyzer views">
              {visibleTabs.map((tab) => {
                const Icon = tab.icon
                return (
                  <button key={tab.id} className={activeTab === tab.id ? 'active' : ''} onClick={() => setActiveTab(tab.id)}>
                    <Icon size={16} />
                    {tab.label}
                  </button>
                )
              })}
            </nav>

            {activeTab === 'overview' && <Overview analysis={analysis} eventCounts={eventCounts} />}
            {activeTab === 'trends' && <Trends analysis={analysis} />}
            {activeTab === 'events' && <Events events={analysis.events} />}
            {activeTab === 'parameters' && <Parameters parameters={analysis.parameterStats || []} />}
            {activeTab === 'raw' && (
              <RawTabs analysis={analysis} selectedRawSheet={selectedRawSheet} onSelectSheet={setSelectedRawSheet} />
            )}
          </>
        ) : (
          <EmptyState />
        )}
      </main>
    </div>
  )
}

function Overview({ analysis, eventCounts }) {
  const stats = analysis.stats
  return (
    <div className="view-stack">
      <section className="summary-grid">
        <Metric title="Samples" value={formatCompact(stats.sampleCount, 1)} detail={analysis.selectedSheet} />
        <Metric title="Log Duration" value={durationLabel(stats.durationMs)} detail={`${formatDateTime(stats.startTime)} to ${formatDateTime(stats.endTime)}`} />
        <Metric title="Run Duration" value={durationLabel(stats.runDurationMs)} detail={`${formatNumber((stats.runningSamples / Math.max(stats.sampleCount, 1)) * 100, 1)}% running samples`} />
        <Metric title="Events" value={formatNumber(eventCounts.critical + eventCounts.warning, 0)} detail={`${eventCounts.critical} critical, ${eventCounts.warning} warnings`} tone={eventCounts.critical ? 'danger' : eventCounts.warning ? 'warn' : 'ok'} />
      </section>

      <section className="overview-layout">
        <div className="panel">
          <div className="panel-header">
            <h2>Operating Snapshot</h2>
            <span>{analysis.fileName}</span>
          </div>
          <div className="signal-grid">
            <Signal label="Max RPM" stat={stats.bySignal.rpm} field="max" />
            <Signal label="Max Frequency" stat={stats.bySignal.frequency} field="max" />
            <Signal label="Min Running Oil Pressure" value={stats.minOilPressureWhileRunning} unit={stats.bySignal.oilPressure?.unit} />
            <Signal label="Max Coolant Temp" value={stats.maxCoolantTemp} unit="deg" />
            <Signal label="Max Oil Temp" value={stats.maxOilTemp} unit="deg" />
            <Signal label="Max Voltage Imbalance" value={stats.maxVoltageImbalance} unit="%" />
            <Signal label="Max Load" stat={stats.bySignal.calculatedLoad} field="max" />
            <Signal label="Max Turbo Speed" stat={stats.bySignal.turboSpeed} field="max" />
            <Signal label="Min Battery Voltage" stat={stats.bySignal.batteryVoltage} field="min" />
          </div>
        </div>

        <div className="panel">
          <div className="panel-header">
            <h2>Detected Channels</h2>
            <span>{analysis.channels.length} mapped</span>
          </div>
          <div className="channel-list">
            {analysis.channels.map((channel) => (
              <div key={channel.key}>
                <span>{channel.label}</span>
                <strong>{channel.source}</strong>
              </div>
            ))}
          </div>
        </div>
      </section>

      <section className="panel">
        <div className="panel-header">
          <h2>Analyzed Samples</h2>
          <span>{analysis.samples.length.toLocaleString()} rows</span>
        </div>
        <SampleTable rows={analysis.samples} />
      </section>
    </div>
  )
}

function Trends({ analysis }) {
  const chartLabel = analysis.chartMeta?.downsampled
    ? `${analysis.chartMeta.chartPointCount.toLocaleString()} rendered from ${analysis.chartMeta.fullSampleCount.toLocaleString()} samples`
    : `All ${analysis.chartMeta?.fullSampleCount?.toLocaleString() || analysis.chartData.length.toLocaleString()} samples`

  return (
    <div className="view-stack">
      {chartGroups.map((group) => (
        <section className="panel chart-panel" key={group.title}>
          <div className="panel-header">
            <h2>{group.title}</h2>
            <span>{chartLabel}</span>
          </div>
          <ResponsiveContainer width="100%" height={310}>
            <LineChart data={analysis.chartData} margin={{ top: 8, right: 20, left: 0, bottom: 8 }}>
              <CartesianGrid strokeDasharray="3 3" stroke="#d8dee8" />
              <XAxis dataKey="timeLabel" minTickGap={36} tick={{ fontSize: 12 }} />
              <YAxis tick={{ fontSize: 12 }} width={56} />
              <Tooltip labelFormatter={(label) => `Time ${label}`} />
              <Legend />
              {group.lines.map((line) => (
                <Line
                  key={line.key}
                  type="monotone"
                  dataKey={line.key}
                  name={line.name}
                  stroke={line.color}
                  strokeWidth={2}
                  dot={false}
                  connectNulls={false}
                  isAnimationActive={false}
                />
              ))}
            </LineChart>
          </ResponsiveContainer>
        </section>
      ))}
    </div>
  )
}

function Events({ events }) {
  return (
    <section className="panel">
      <div className="panel-header">
        <h2>Detected Events</h2>
        <span>{events.length} grouped events</span>
      </div>
      <div className="event-list">
        {events.map((event, index) => (
          <div className={`event-row ${event.severity}`} key={`${event.title}-${event.time}-${index}`}>
            <div>
              <span className="severity">{event.severity}</span>
              <strong>{event.title}</strong>
              <p>{event.detail}</p>
            </div>
            <div className="event-meta">
              <span>{formatDateTime(event.time)}</span>
              {event.rowNumber ? <span>Row {event.rowNumber}</span> : null}
            </div>
          </div>
        ))}
      </div>
    </section>
  )
}

function Parameters({ parameters }) {
  const [page, setPage] = useState(0)
  const [pageSize, setPageSize] = useState(100)
  const [query, setQuery] = useState('')
  const filtered = useMemo(() => {
    const needle = query.trim().toLowerCase()
    if (!needle) return parameters
    return parameters.filter((parameter) =>
      `${parameter.header} ${parameter.mappedSignal} ${parameter.type}`.toLowerCase().includes(needle),
    )
  }, [parameters, query])
  const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize))
  const safePage = Math.min(page, totalPages - 1)
  const visibleRows = filtered.slice(safePage * pageSize, safePage * pageSize + pageSize)

  useEffect(() => {
    setPage(0)
  }, [query, pageSize])

  return (
    <section className="panel">
      <div className="panel-header">
        <h2>Parameter Summary</h2>
        <span>{filtered.length.toLocaleString()} of {parameters.length.toLocaleString()} parameters</span>
      </div>
      <div className="filter-row">
        <input value={query} onChange={(event) => setQuery(event.target.value)} placeholder="Filter parameters" />
      </div>
      <Pagination
        page={safePage}
        pageSize={pageSize}
        totalRows={filtered.length}
        totalPages={totalPages}
        onPageChange={setPage}
        onPageSizeChange={setPageSize}
      />
      <div className="sample-table-wrap">
        <table className="sample-table parameter-table">
          <thead>
            <tr>
              <th>#</th>
              <th>Parameter</th>
              <th>Mapped Signal</th>
              <th>Type</th>
              <th>Count</th>
              <th>Min</th>
              <th>Avg</th>
              <th>Max</th>
              <th>Text Samples</th>
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((parameter) => (
              <tr key={`${parameter.index}-${parameter.header}`}>
                <td>{parameter.index + 1}</td>
                <td>{parameter.header}</td>
                <td>{parameter.mappedSignal}</td>
                <td>{parameter.type}</td>
                <td>{parameter.count.toLocaleString()}</td>
                <td>{formatNumber(parameter.min, 2)}</td>
                <td>{formatNumber(parameter.avg, 2)}</td>
                <td>{formatNumber(parameter.max, 2)}</td>
                <td>{parameter.samples?.join(', ') || ''}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </section>
  )
}

function RawTabs({ analysis, selectedRawSheet, onSelectSheet }) {
  const sheet = analysis.sheets.find((item) => item.name === selectedRawSheet) || analysis.sheets[0]
  const [page, setPage] = useState(0)
  const [pageSize, setPageSize] = useState(100)
  const rows = sheet?.tableRows || sheet?.sampleRows || []
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize))
  const safePage = Math.min(page, totalPages - 1)
  const visibleRows = rows.slice(safePage * pageSize, safePage * pageSize + pageSize)

  useEffect(() => {
    setPage(0)
  }, [sheet?.name, pageSize])

  return (
    <div className="raw-layout">
      <aside className="sheet-list">
        {analysis.sheets.map((item) => (
          <button key={item.name} className={item.name === sheet.name ? 'active' : ''} onClick={() => onSelectSheet(item.name)}>
            <span>{item.name}</span>
            <small>{item.rows.toLocaleString()} rows x {item.cols} cols</small>
          </button>
        ))}
      </aside>
      <section className="panel raw-panel">
        <div className="panel-header">
          <h2>{sheet.name}</h2>
          <span>{rows.length.toLocaleString()} visible rows available</span>
        </div>
        <Pagination
          page={safePage}
          pageSize={pageSize}
          totalRows={rows.length}
          totalPages={totalPages}
          onPageChange={setPage}
          onPageSizeChange={setPageSize}
        />
        <div className="raw-table-wrap">
          <table className="raw-table">
            <tbody>
              {visibleRows.map((row, rowIndex) => (
                <tr key={`${safePage}-${rowIndex}`}>
                  {Array.from({ length: sheet.cols }).map((_, cellIndex) => (
                    <td key={cellIndex}>{row[cellIndex] ?? ''}</td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </section>
    </div>
  )
}

function Pagination({ page, pageSize, totalRows, totalPages, onPageChange, onPageSizeChange }) {
  const start = totalRows === 0 ? 0 : page * pageSize + 1
  const end = Math.min(totalRows, (page + 1) * pageSize)
  return (
    <div className="pagination">
      <span>{start.toLocaleString()}-{end.toLocaleString()} of {totalRows.toLocaleString()}</span>
      <div>
        <button onClick={() => onPageChange(0)} disabled={page === 0}>First</button>
        <button onClick={() => onPageChange(Math.max(0, page - 1))} disabled={page === 0}>Prev</button>
        <button onClick={() => onPageChange(Math.min(totalPages - 1, page + 1))} disabled={page >= totalPages - 1}>Next</button>
        <button onClick={() => onPageChange(totalPages - 1)} disabled={page >= totalPages - 1}>Last</button>
        <select value={pageSize} onChange={(event) => onPageSizeChange(Number(event.target.value))}>
          <option value={50}>50 rows</option>
          <option value={100}>100 rows</option>
          <option value={250}>250 rows</option>
          <option value={500}>500 rows</option>
          <option value={2000}>2,000 rows</option>
        </select>
      </div>
    </div>
  )
}

function EmptyState() {
  return (
    <section className="empty-state">
      <Zap size={30} />
      <h2>Ready for the first Gen-Set log</h2>
      <p>The analyzer will identify the timestamped data tab, map common engine and generator channels, remove invalid sentinel values, and build a dashboard for review.</p>
    </section>
  )
}

function Metric({ title, value, detail, tone }) {
  return (
    <div className={`metric ${tone || ''}`}>
      <span>{title}</span>
      <strong>{value}</strong>
      <p>{detail}</p>
    </div>
  )
}

function Signal({ label, stat, field, value, unit }) {
  const displayValue = value ?? stat?.[field]
  const displayUnit = unit ?? stat?.unit ?? ''
  return (
    <div className="signal">
      <span>{label}</span>
      <strong>{formatNumber(displayValue, displayUnit === '%' ? 2 : 1)}{displayUnit ? ` ${displayUnit}` : ''}</strong>
    </div>
  )
}

function SampleTable({ rows }) {
  const [page, setPage] = useState(0)
  const [pageSize, setPageSize] = useState(2000)
  const totalPages = Math.max(1, Math.ceil(rows.length / pageSize))
  const safePage = Math.min(page, totalPages - 1)
  const visibleRows = rows.slice(safePage * pageSize, safePage * pageSize + pageSize)

  useEffect(() => {
    setPage(0)
  }, [rows, pageSize])

  return (
    <>
      <Pagination
        page={safePage}
        pageSize={pageSize}
        totalRows={rows.length}
        totalPages={totalPages}
        onPageChange={setPage}
        onPageSizeChange={setPageSize}
      />
      <div className="sample-table-wrap">
        <table className="sample-table">
          <thead>
            <tr>
              <th>Row</th>
              <th>Time</th>
              <th>Elapsed</th>
              <th>RPM</th>
              <th>Hz</th>
              <th>Oil Press.</th>
              <th>Coolant</th>
              <th>Oil Temp</th>
              <th>Load</th>
              <th>Battery</th>
              <th>Rail Press.</th>
              <th>Volts Avg</th>
              <th>State</th>
              <th>Operation</th>
            </tr>
          </thead>
          <tbody>
            {visibleRows.map((row) => (
              <tr key={`${row.rowNumber}-${row.isoTime}`}>
                <td>{row.rowNumber}</td>
                <td>{formatDateTime(row.isoTime)}</td>
                <td>{formatNumber(row.elapsedSeconds, 1)}s</td>
                <td>{formatNumber(row.rpm, 0)}</td>
                <td>{formatNumber(row.frequency, 2)}</td>
                <td>{formatNumber(row.oilPressure, 1)}</td>
                <td>{formatNumber(row.coolantTemp, 1)}</td>
                <td>{formatNumber(row.oilTemp, 1)}</td>
                <td>{formatNumber(row.calculatedLoad, 1)}</td>
                <td>{formatNumber(row.batteryVoltage, 2)}</td>
                <td>{formatNumber(row.railPressure, 0)}</td>
                <td>{formatNumber(row.voltageAvg, 1)}</td>
                <td>{formatNumber(row.engineState, 0)}</td>
                <td>{row.engineOperation || ''}</td>
              </tr>
            ))}
          </tbody>
        </table>
      </div>
    </>
  )
}

export default App
