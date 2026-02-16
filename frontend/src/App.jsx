import { useState, useEffect, useCallback } from 'react'

const API_BASE_URL = import.meta.env.VITE_API_URL || 'http://localhost:8000/api'

const FILE_TYPES = [
  { key: 'aramco', label: 'Saudi Aramco 9COM', color: 'bg-green-600', accept: '.xlsx,.xls' },
  { key: 'cfihos', label: 'CFIHOS Standard', color: 'bg-blue-600', accept: '.xlsx,.xls' },
  { key: 'kbr', label: 'KBR FEED', color: 'bg-red-600', accept: '.xlsx,.xls' },
  { key: 'ltc', label: 'LTC Contractor', color: 'bg-purple-600', accept: '.xlsx,.xls' },
  { key: 'sa_doc', label: 'SA Document', color: 'bg-amber-600', accept: '.xlsx,.xls' },
]

// ─── Tabs ──────────────────────────────────────────────────────────────
const TABS = ['Dashboard', 'Upload & Configure', 'Search', 'Batch', 'Logs']

export default function App() {
  const [tab, setTab] = useState('Upload & Configure')
  const [stats, setStats] = useState(null)
  const [classes, setClasses] = useState([])
  const [masterConfig, setMasterConfig] = useState({ masters: [], available: [] })
  const [logs, setLogs] = useState([])
  const [searchQuery, setSearchQuery] = useState('')
  const [searchResults, setSearchResults] = useState([])
  const [batchText, setBatchText] = useState('')
  const [batchResults, setBatchResults] = useState([])
  const [uploads, setUploads] = useState({})
  const [loading, setLoading] = useState(false)
  const [selectedClass, setSelectedClass] = useState(null)
  const [pendingMasters, setPendingMasters] = useState([])

  // ─── Fetch helpers ──────────────────────────────────────────────────
  const api = useCallback(async (path, opts = {}) => {
    const res = await fetch(`${API_BASE_URL}${path}`, {
      headers: { 'Content-Type': 'application/json', ...opts.headers },
      ...opts,
    })
    if (!res.ok) {
      const err = await res.json().catch(() => ({ detail: res.statusText }))
      throw new Error(err.detail || 'API error')
    }
    return res.json()
  }, [])

  const refreshAll = useCallback(async () => {
    try {
      const [s, mc, c] = await Promise.all([
        api('/stats'),
        api('/master-config'),
        api('/classes?limit=500'),
      ])
      setStats(s)
      setMasterConfig(mc)
      setPendingMasters(mc.masters)
      setClasses(c.items || [])
    } catch { /* ignore on initial load */ }
  }, [api])

  useEffect(() => { refreshAll() }, [refreshAll])

  // ─── Upload ─────────────────────────────────────────────────────────
  const handleUpload = async (fileType, file) => {
    const fd = new FormData()
    fd.append('file', file)
    setLoading(true)
    try {
      const res = await fetch(`${API_BASE_URL}/upload/${fileType}`, { method: 'POST', body: fd })
      const data = await res.json()
      if (!res.ok) throw new Error(data.detail || 'Upload failed')
      setUploads(u => ({ ...u, [fileType]: data }))
      await refreshAll()
    } catch (e) {
      alert(e.message)
    } finally {
      setLoading(false)
    }
  }

  // ─── Master config ──────────────────────────────────────────────────
  const togglePendingMaster = (key) => {
    setPendingMasters(prev => {
      if (prev.includes(key)) return prev.filter(k => k !== key)
      if (prev.length >= 2) return [prev[1], key]   // rotate: keep last, add new
      return [...prev, key]
    })
  }

  const saveMasterConfig = async () => {
    if (pendingMasters.length === 0) return alert('Select at least 1 master file.')
    setLoading(true)
    try {
      await api('/master-config', {
        method: 'POST',
        body: JSON.stringify({ masters: pendingMasters }),
      })
      await refreshAll()
    } catch (e) {
      alert(e.message)
    } finally {
      setLoading(false)
    }
  }

  // ─── Harmonize ──────────────────────────────────────────────────────
  const runHarmonize = async () => {
    setLoading(true)
    try {
      await api('/harmonize', { method: 'POST' })
      await refreshAll()
      setTab('Dashboard')
    } catch (e) {
      alert(e.message)
    } finally {
      setLoading(false)
    }
  }

  // ─── Demo ───────────────────────────────────────────────────────────
  const loadDemo = async () => {
    setLoading(true)
    try {
      const params = pendingMasters.length > 0
        ? '?' + pendingMasters.map(m => `masters=${m}`).join('&')
        : '?masters=aramco'
      await api(`/load-demo${params}`, { method: 'POST' })
      await refreshAll()
      setTab('Dashboard')
    } catch (e) {
      alert(e.message)
    } finally {
      setLoading(false)
    }
  }

  // ─── Search ─────────────────────────────────────────────────────────
  const runSearch = async () => {
    if (!searchQuery.trim()) return
    setLoading(true)
    try {
      const data = await api('/search', {
        method: 'POST',
        body: JSON.stringify({ query: searchQuery, limit: 20 }),
      })
      setSearchResults(data.results || [])
    } catch (e) {
      alert(e.message)
    } finally {
      setLoading(false)
    }
  }

  // ─── Batch ──────────────────────────────────────────────────────────
  const runBatch = async () => {
    const items = batchText.split('\n').map(s => s.trim()).filter(Boolean)
    if (!items.length) return
    setLoading(true)
    try {
      const data = await api('/batch', {
        method: 'POST',
        body: JSON.stringify({ items }),
      })
      setBatchResults(data.results || [])
    } catch (e) {
      alert(e.message)
    } finally {
      setLoading(false)
    }
  }

  // ─── Export ─────────────────────────────────────────────────────────
  const exportCSV = () => {
    window.open(`${API_BASE_URL}/export/csv`, '_blank')
  }

  // ─── Logs ───────────────────────────────────────────────────────────
  const fetchLogs = async () => {
    try {
      const data = await api('/logs')
      setLogs(data.logs || [])
    } catch { /* ignore */ }
  }
  useEffect(() => { if (tab === 'Logs') fetchLogs() }, [tab])

  // ─── Score badge ────────────────────────────────────────────────────
  const ScoreBadge = ({ score }) => {
    const color = score >= 90 ? 'bg-green-100 text-green-800'
      : score >= 70 ? 'bg-yellow-100 text-yellow-800'
      : 'bg-red-100 text-red-800'
    return <span className={`px-2 py-0.5 rounded text-xs font-semibold ${color}`}>{score}%</span>
  }

  // ═══════════════════════════════════════════════════════════════════
  // RENDER
  // ═══════════════════════════════════════════════════════════════════
  return (
    <div className="min-h-screen bg-gray-50">
      {/* Header */}
      <header className="bg-gradient-to-r from-red-700 to-red-900 text-white shadow-lg">
        <div className="max-w-7xl mx-auto px-4 py-4 flex items-center justify-between">
          <div>
            <h1 className="text-2xl font-bold tracking-tight">KBR RDL Data Harmonizer</h1>
            <p className="text-red-200 text-sm">Equipment Class Harmonization v2.0</p>
          </div>
          <div className="flex items-center gap-3">
            {stats && (
              <span className="text-sm bg-white/20 px-3 py-1 rounded-full">
                {stats.total} classes
              </span>
            )}
            {stats?.masters?.length > 0 && (
              <span className="text-sm bg-white/20 px-3 py-1 rounded-full">
                Master: {stats.masters.join(' + ')}
              </span>
            )}
          </div>
        </div>
      </header>

      {/* Tab bar */}
      <nav className="bg-white border-b shadow-sm">
        <div className="max-w-7xl mx-auto px-4 flex gap-1">
          {TABS.map(t => (
            <button
              key={t}
              onClick={() => setTab(t)}
              className={`px-4 py-3 text-sm font-medium border-b-2 transition-colors ${
                tab === t
                  ? 'border-red-600 text-red-700'
                  : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
              }`}
            >
              {t}
            </button>
          ))}
        </div>
      </nav>

      <main className="max-w-7xl mx-auto px-4 py-6">
        {loading && (
          <div className="fixed inset-0 bg-black/20 flex items-center justify-center z-50">
            <div className="bg-white rounded-lg p-6 shadow-xl flex items-center gap-3">
              <svg className="animate-spin h-6 w-6 text-red-600" viewBox="0 0 24 24">
                <circle className="opacity-25" cx="12" cy="12" r="10" stroke="currentColor" strokeWidth="4" fill="none"/>
                <path className="opacity-75" fill="currentColor" d="M4 12a8 8 0 018-8v4a4 4 0 00-4 4H4z"/>
              </svg>
              Processing...
            </div>
          </div>
        )}

        {/* ─── UPLOAD & CONFIGURE ─────────────────────────────────── */}
        {tab === 'Upload & Configure' && (
          <div className="space-y-6">
            {/* Master Selection */}
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-lg font-semibold mb-1">Choose Master File(s)</h2>
              <p className="text-gray-500 text-sm mb-4">
                Select <strong>1 or 2</strong> file types as the master source. The master defines
                the base set of equipment classes; all other uploaded files are matched against it.
              </p>

              <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-3 mb-4">
                {FILE_TYPES.map(ft => {
                  const isSelected = pendingMasters.includes(ft.key)
                  const isUploaded = masterConfig.available?.includes(ft.key)
                  return (
                    <button
                      key={ft.key}
                      onClick={() => togglePendingMaster(ft.key)}
                      className={`relative rounded-lg border-2 p-4 text-left transition-all ${
                        isSelected
                          ? 'border-red-600 bg-red-50 ring-2 ring-red-200'
                          : 'border-gray-200 hover:border-gray-400'
                      }`}
                    >
                      <div className="flex items-center gap-2">
                        <span className={`w-3 h-3 rounded-full ${ft.color}`} />
                        <span className="font-medium text-sm">{ft.label}</span>
                      </div>
                      {isUploaded && (
                        <span className="absolute top-2 right-2 text-xs text-green-600 font-medium">
                          Uploaded
                        </span>
                      )}
                      {isSelected && (
                        <span className="mt-2 inline-block text-xs bg-red-600 text-white px-2 py-0.5 rounded">
                          MASTER {pendingMasters.indexOf(ft.key) + 1}
                        </span>
                      )}
                    </button>
                  )
                })}
              </div>

              <div className="flex gap-3">
                <button
                  onClick={saveMasterConfig}
                  disabled={pendingMasters.length === 0}
                  className="px-4 py-2 bg-red-600 text-white rounded-lg text-sm font-medium hover:bg-red-700 disabled:opacity-50 disabled:cursor-not-allowed"
                >
                  Save Master Config
                </button>
                <span className="text-sm text-gray-400 self-center">
                  {pendingMasters.length === 0
                    ? 'Click a card to select a master'
                    : `Selected: ${pendingMasters.join(', ')}`}
                </span>
              </div>
            </div>

            {/* File Upload */}
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="text-lg font-semibold mb-4">Upload Data Files</h2>
              <div className="grid sm:grid-cols-2 lg:grid-cols-3 gap-4">
                {FILE_TYPES.map(ft => (
                  <div key={ft.key} className="border rounded-lg p-4">
                    <div className="flex items-center gap-2 mb-2">
                      <span className={`w-3 h-3 rounded-full ${ft.color}`} />
                      <span className="font-medium text-sm">{ft.label}</span>
                      {masterConfig.masters?.includes(ft.key) && (
                        <span className="text-xs bg-red-100 text-red-700 px-1.5 py-0.5 rounded font-medium">
                          Master
                        </span>
                      )}
                    </div>
                    {uploads[ft.key] ? (
                      <p className="text-xs text-green-600">
                        {uploads[ft.key].filename} — {uploads[ft.key].records} records
                      </p>
                    ) : (
                      <p className="text-xs text-gray-400 mb-2">No file uploaded</p>
                    )}
                    <label className="inline-block mt-2 px-3 py-1.5 bg-gray-100 hover:bg-gray-200 rounded text-xs font-medium cursor-pointer">
                      Choose file
                      <input
                        type="file"
                        className="hidden"
                        accept={ft.accept}
                        onChange={e => e.target.files[0] && handleUpload(ft.key, e.target.files[0])}
                      />
                    </label>
                  </div>
                ))}
              </div>
            </div>

            {/* Actions */}
            <div className="flex gap-3">
              <button
                onClick={runHarmonize}
                className="px-6 py-2.5 bg-red-600 text-white rounded-lg font-medium hover:bg-red-700 shadow"
              >
                Run Harmonization
              </button>
              <button
                onClick={loadDemo}
                className="px-6 py-2.5 bg-gray-600 text-white rounded-lg font-medium hover:bg-gray-700"
              >
                Load Demo Data
              </button>
              <button
                onClick={exportCSV}
                className="px-6 py-2.5 border border-gray-300 rounded-lg font-medium hover:bg-gray-50"
              >
                Export CSV
              </button>
            </div>
          </div>
        )}

        {/* ─── DASHBOARD ──────────────────────────────────────────── */}
        {tab === 'Dashboard' && (
          <div className="space-y-6">
            {/* Metrics */}
            {stats && (
              <div className="grid sm:grid-cols-2 lg:grid-cols-4 gap-4">
                <MetricCard
                  label="Total Classes"
                  value={stats.total}
                  sub={`Master: ${stats.masters?.join(' + ') || '—'}`}
                />
                {Object.entries(stats.matches || {}).map(([src, pct]) => (
                  <MetricCard key={src} label={`${src.toUpperCase()} Match`} value={`${pct}%`} sub={`match rate`} />
                ))}
                <MetricCard label="Gaps" value={stats.gap_count || 0} sub="missing links" />
              </div>
            )}

            {/* Class Table */}
            <div className="bg-white rounded-lg shadow overflow-hidden">
              <div className="px-6 py-4 border-b flex items-center justify-between">
                <h2 className="font-semibold">Harmonized Equipment Classes</h2>
                <span className="text-sm text-gray-500">{classes.length} items</span>
              </div>
              <div className="overflow-x-auto">
                <table className="w-full text-sm">
                  <thead className="bg-gray-50 text-left">
                    <tr>
                      <th className="px-4 py-3 font-medium">#</th>
                      <th className="px-4 py-3 font-medium">Master Name</th>
                      <th className="px-4 py-3 font-medium">Source</th>
                      <th className="px-4 py-3 font-medium">Matches</th>
                      <th className="px-4 py-3 font-medium">Gaps</th>
                      <th className="px-4 py-3 font-medium"></th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {classes.map(c => (
                      <tr key={c.uid} className="hover:bg-gray-50">
                        <td className="px-4 py-3 text-gray-500">{c.index}</td>
                        <td className="px-4 py-3 font-medium">{c.master_name}</td>
                        <td className="px-4 py-3">
                          <span className="px-2 py-0.5 rounded text-xs bg-gray-100 font-medium">
                            {c.master_source}
                          </span>
                        </td>
                        <td className="px-4 py-3">
                          <div className="flex gap-1 flex-wrap">
                            {Object.entries(c.matches).map(([src, m]) => (
                              <span key={src} className="inline-flex items-center gap-1">
                                <span className="text-xs text-gray-600">{src}:</span>
                                <ScoreBadge score={m.score} />
                              </span>
                            ))}
                          </div>
                        </td>
                        <td className="px-4 py-3">
                          {c.gaps.length > 0 && (
                            <span className="text-xs text-red-600 font-medium">
                              {c.gaps.join(', ')}
                            </span>
                          )}
                        </td>
                        <td className="px-4 py-3">
                          <button
                            onClick={() => setSelectedClass(selectedClass?.uid === c.uid ? null : c)}
                            className="text-xs text-red-600 hover:underline"
                          >
                            {selectedClass?.uid === c.uid ? 'Hide' : 'Details'}
                          </button>
                        </td>
                      </tr>
                    ))}
                    {classes.length === 0 && (
                      <tr>
                        <td colSpan={6} className="px-4 py-8 text-center text-gray-400">
                          No data yet. Upload files and run harmonization, or load demo data.
                        </td>
                      </tr>
                    )}
                  </tbody>
                </table>
              </div>
            </div>

            {/* Detail panel */}
            {selectedClass && (
              <div className="bg-white rounded-lg shadow p-6">
                <h3 className="font-semibold text-lg mb-4">
                  {selectedClass.master_name}
                  <span className="ml-2 text-sm text-gray-400">({selectedClass.master_source})</span>
                </h3>
                <div className="grid md:grid-cols-2 gap-6">
                  {/* Matches */}
                  <div>
                    <h4 className="font-medium text-sm text-gray-600 mb-2">System Matches</h4>
                    <div className="space-y-2">
                      {Object.entries(selectedClass.matches).map(([src, m]) => (
                        <div key={src} className="flex items-center justify-between bg-gray-50 rounded p-3">
                          <div>
                            <span className="text-xs font-semibold uppercase text-gray-500">{src}</span>
                            <p className="font-medium text-sm">{m.name}</p>
                            <p className="text-xs text-gray-400">ID: {m.id}</p>
                          </div>
                          <ScoreBadge score={m.score} />
                        </div>
                      ))}
                    </div>
                  </div>
                  {/* Gaps */}
                  <div>
                    <h4 className="font-medium text-sm text-gray-600 mb-2">Gaps</h4>
                    {selectedClass.gaps.length > 0 ? (
                      <ul className="space-y-1">
                        {selectedClass.gaps.map(g => (
                          <li key={g} className="flex items-center gap-2 text-sm text-red-600">
                            <span className="w-2 h-2 bg-red-400 rounded-full" /> No match in {g}
                          </li>
                        ))}
                      </ul>
                    ) : (
                      <p className="text-sm text-green-600">All systems matched.</p>
                    )}
                  </div>
                </div>
              </div>
            )}
          </div>
        )}

        {/* ─── SEARCH ─────────────────────────────────────────────── */}
        {tab === 'Search' && (
          <div className="space-y-4">
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="font-semibold mb-3">Search Equipment Classes</h2>
              <div className="flex gap-2">
                <input
                  value={searchQuery}
                  onChange={e => setSearchQuery(e.target.value)}
                  onKeyDown={e => e.key === 'Enter' && runSearch()}
                  placeholder="e.g. centrifugal pump, heat exchanger..."
                  className="flex-1 border rounded-lg px-4 py-2 text-sm focus:ring-2 focus:ring-red-300 focus:outline-none"
                />
                <button onClick={runSearch} className="px-5 py-2 bg-red-600 text-white rounded-lg text-sm font-medium hover:bg-red-700">
                  Search
                </button>
              </div>
            </div>
            {searchResults.length > 0 && (
              <div className="bg-white rounded-lg shadow overflow-hidden">
                <table className="w-full text-sm">
                  <thead className="bg-gray-50 text-left">
                    <tr>
                      <th className="px-4 py-3 font-medium">Name</th>
                      <th className="px-4 py-3 font-medium">Source</th>
                      <th className="px-4 py-3 font-medium">Score</th>
                      <th className="px-4 py-3 font-medium">Matches</th>
                      <th className="px-4 py-3 font-medium">Gaps</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {searchResults.map(r => (
                      <tr key={r.uid} className="hover:bg-gray-50">
                        <td className="px-4 py-3 font-medium">{r.master_name}</td>
                        <td className="px-4 py-3 text-xs">{r.master_source}</td>
                        <td className="px-4 py-3"><ScoreBadge score={r.search_score} /></td>
                        <td className="px-4 py-3 text-xs">{Object.keys(r.matches).join(', ')}</td>
                        <td className="px-4 py-3 text-xs text-red-600">{r.gaps.join(', ')}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* ─── BATCH ──────────────────────────────────────────────── */}
        {tab === 'Batch' && (
          <div className="space-y-4">
            <div className="bg-white rounded-lg shadow p-6">
              <h2 className="font-semibold mb-3">Batch Process</h2>
              <p className="text-sm text-gray-500 mb-3">Enter one equipment name per line.</p>
              <textarea
                value={batchText}
                onChange={e => setBatchText(e.target.value)}
                rows={6}
                placeholder={"Centrifugal Pump\nHeat Exchanger\nPressure Vessel"}
                className="w-full border rounded-lg px-4 py-2 text-sm focus:ring-2 focus:ring-red-300 focus:outline-none"
              />
              <button onClick={runBatch} className="mt-3 px-5 py-2 bg-red-600 text-white rounded-lg text-sm font-medium hover:bg-red-700">
                Process Batch
              </button>
            </div>
            {batchResults.length > 0 && (
              <div className="bg-white rounded-lg shadow overflow-hidden">
                <table className="w-full text-sm">
                  <thead className="bg-gray-50 text-left">
                    <tr>
                      <th className="px-4 py-3 font-medium">Input</th>
                      <th className="px-4 py-3 font-medium">Best Match</th>
                      <th className="px-4 py-3 font-medium">Score</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y">
                    {batchResults.map((r, i) => (
                      <tr key={i} className="hover:bg-gray-50">
                        <td className="px-4 py-3">{r.input}</td>
                        <td className="px-4 py-3 font-medium">{r.match?.master_name || '—'}</td>
                        <td className="px-4 py-3">{r.score > 0 ? <ScoreBadge score={r.score} /> : '—'}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </div>
        )}

        {/* ─── LOGS ───────────────────────────────────────────────── */}
        {tab === 'Logs' && (
          <div className="bg-white rounded-lg shadow p-6">
            <div className="flex items-center justify-between mb-4">
              <h2 className="font-semibold">System Logs</h2>
              <button onClick={fetchLogs} className="text-sm text-red-600 hover:underline">Refresh</button>
            </div>
            <div className="bg-gray-900 text-gray-200 rounded-lg p-4 font-mono text-xs max-h-96 overflow-y-auto space-y-1">
              {logs.length === 0 && <p className="text-gray-500">No logs yet.</p>}
              {logs.map((l, i) => (
                <p key={i}>
                  <span className="text-gray-500">{l.ts}</span>{' '}
                  <span className={l.level === 'error' ? 'text-red-400' : l.level === 'warning' ? 'text-yellow-400' : 'text-green-400'}>
                    [{l.level}]
                  </span>{' '}
                  {l.msg}
                </p>
              ))}
            </div>
          </div>
        )}
      </main>
    </div>
  )
}

function MetricCard({ label, value, sub }) {
  return (
    <div className="bg-white rounded-lg shadow p-5">
      <p className="text-sm text-gray-500">{label}</p>
      <p className="text-3xl font-bold text-gray-800 mt-1">{value}</p>
      {sub && <p className="text-xs text-gray-400 mt-1">{sub}</p>}
    </div>
  )
}
