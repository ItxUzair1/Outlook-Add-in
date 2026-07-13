import { useState, useEffect, useRef, useCallback } from 'react';
import {
  BarChart2, TrendingUp, Clock, Calendar, BarChart,
  RefreshCw, Layers, Activity, ChevronDown
} from 'lucide-react';

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:4001/api';
const POLL_MS = 2500;

// ─── tiny pie / donut chart ──────────────────────────────────────────────────
function DonutChart({ slices, size = 100 }) {
  if (!slices || slices.length === 0) return null;
  const total = slices.reduce((s, sl) => s + sl.value, 0);
  if (total === 0) return null;

  const cx = size / 2, cy = size / 2, r = size / 2 - 6, ri = r * 0.55;
  const angles = slices.reduce((acc, sl) => {
    const prev = acc[acc.length - 1];
    acc.push(prev + (sl.value / total) * 2 * Math.PI);
    return acc;
  }, [-Math.PI / 2]);

  const paths = slices.map((sl, i) => {
    const startAngle = angles[i];
    const endAngle = angles[i + 1];
    const sweep = endAngle - startAngle;
    const x1o = cx + r * Math.cos(startAngle);
    const y1o = cy + r * Math.sin(startAngle);
    const x1i = cx + ri * Math.cos(startAngle);
    const y1i = cy + ri * Math.sin(startAngle);
    const x2o = cx + r * Math.cos(endAngle);
    const y2o = cy + r * Math.sin(endAngle);
    const x2i = cx + ri * Math.cos(endAngle);
    const y2i = cy + ri * Math.sin(endAngle);
    const large = sweep > Math.PI ? 1 : 0;
    const d = `M${x1o},${y1o} A${r},${r} 0 ${large},1 ${x2o},${y2o} L${x2i},${y2i} A${ri},${ri} 0 ${large},0 ${x1i},${y1i} Z`;
    
    const pct = Math.round((sl.value / total) * 100);
    const midAngle = startAngle + sweep / 2;
    const textR = r * 0.77;
    const tx = cx + textR * Math.cos(midAngle);
    const ty = cy + textR * Math.sin(midAngle);

    return (
      <g key={i}>
        <path d={d} fill={sl.color}>
          <title>{`${sl.label}: ${sl.value}`}</title>
        </path>
        {pct >= 5 && (
          <text x={tx} y={ty} fill="#ffffff" fontSize={size * 0.07} fontWeight="bold" textAnchor="middle" dominantBaseline="central" pointerEvents="none">
            {pct}%
          </text>
        )}
      </g>
    );
  });

  return (
    <div style={{ position: 'relative', width: size, height: size }}>
      <svg width={size} height={size} style={{ flexShrink: 0 }}>
        {paths}
      </svg>
    </div>
  );
}

// ─── KPI Card ────────────────────────────────────────────────────────────────
function KpiCard({ label, value, icon: Icon, color, subtitle }) {
  return (
    <div className="an-kpi-card" style={{ '--kpi-color': color }}>
      <div className="an-kpi-icon" style={{ background: color + '18', color }}>
        <Icon size={20} />
      </div>
      <div className="an-kpi-body">
        <div className="an-kpi-value">{value ?? '—'}</div>
        <div className="an-kpi-label">{label}</div>
        {subtitle && <div className="an-kpi-sub">{subtitle}</div>}
      </div>
    </div>
  );
}

const FilterSelect = ({ value, onChange, options }) => (
  <select 
    value={value} 
    onChange={e => onChange(e.target.value)}
    style={{ padding: '4px 24px 4px 10px', fontSize: 12, borderRadius: 6, border: '1px solid #e2e8f0', background: '#f8fafc', color: '#64748b', outline: 'none', cursor: 'pointer', appearance: 'auto' }}
  >
    {options.map(opt => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
  </select>
);

// ─── Analytics Page ────────────────────────────────────────────────────────────
export default function Analytics() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [activeWindow, setActiveWindow] = useState('day');
  const [lastFetch, setLastFetch] = useState(null);
  const intervalRef = useRef(null);

  // Panel States
  const [topSearchFilter, setTopSearchFilter] = useState('all');
  const [leastSearchFilter, setLeastSearchFilter] = useState('all');
  const [pieChartFilter, setPieChartFilter] = useState('all');
  
  const [topLimit, setTopLimit] = useState(10);
  const [leastLimit, setLeastLimit] = useState(10);

  const fetchAnalytics = useCallback(async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/analytics`);
      if (resp.ok) {
        const json = await resp.json();
        setData(json);
        setLastFetch(new Date());
      }
    } catch {
      // silently fail
    }
    setLoading(false);
  }, []);

  useEffect(() => {
    const timerId = setTimeout(fetchAnalytics, 0);
    intervalRef.current = setInterval(fetchAnalytics, POLL_MS);
    return () => {
      clearTimeout(timerId);
      clearInterval(intervalRef.current);
    };
  }, [fetchAnalytics]);

  const COLORS = [
    '#0078d4', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#06b6d4', '#ec4899', '#84cc16',
    '#6366f1', '#14b8a6', '#f97316', '#d946ef', '#0ea5e9', '#8b5cf6', '#f43f5e', '#a855f7',
    '#22c55e', '#eab308', '#3b82f6', '#10b981', '#f43f5e', '#64748b'
  ];

  const getAggregations = (filterName) => {
    const projectAgg = {};
    const yearAgg = {};
    const projectYearMap = {};

    if (!data) return { projectAgg, yearAgg, projectYearMap };

    if (filterName === 'all') {
      for (const [year, projects] of Object.entries(data.totals)) {
        let yearSum = 0;
        for (const [proj, cnt] of Object.entries(projects)) {
          yearSum += cnt;
          projectAgg[proj] = (projectAgg[proj] || 0) + cnt;
          if (!projectYearMap[proj] || cnt > (data.totals[projectYearMap[proj]][proj] || 0)) {
              projectYearMap[proj] = year;
          }
        }
        yearAgg[year] = yearSum;
      }
    } else {
      const now = lastFetch ? lastFetch.getTime() : 0;
      const cutoffs = {
        hour: now - 60 * 60 * 1000,
        day: now - 24 * 60 * 60 * 1000,
        week: now - 7 * 24 * 60 * 60 * 1000,
        month: now - 30 * 24 * 60 * 60 * 1000
      };
      
      const cutoff = cutoffs[filterName];
      const filteredEvents = (data.events || []).filter(e => e.ts >= cutoff);
      
      for (const evt of filteredEvents) {
        const proj = evt.project || 'Unknown';
        const year = evt.year || 'Unknown';
        projectAgg[proj] = (projectAgg[proj] || 0) + 1;
        yearAgg[year] = (yearAgg[year] || 0) + 1;
        
        if (!projectYearMap[proj]) {
          projectYearMap[proj] = { [year]: 1 };
        } else {
          projectYearMap[proj][year] = (projectYearMap[proj][year] || 0) + 1;
        }
      }
      
      // Resolve most common year per project
      for (const proj of Object.keys(projectYearMap)) {
         let maxYear = null;
         let maxCount = 0;
         for (const [yr, c] of Object.entries(projectYearMap[proj])) {
           if (c > maxCount) { maxCount = c; maxYear = yr; }
         }
         projectYearMap[proj] = maxYear;
      }
    }
    return { projectAgg, yearAgg, projectYearMap };
  };

  // Precompute Global all-time mapping for consistent Colors
  const { yearAgg: allTimeYearAgg, projectAgg: allTimeProjectAgg } = getAggregations('all');
  const allTimeSortedYears = Object.entries(allTimeYearAgg).sort((a, b) => b[0].localeCompare(a[0]));
  const yearColorMap = {};
  allTimeSortedYears.forEach(([yr], i) => {
    yearColorMap[yr] = COLORS[i % COLORS.length];
  });
  
  let totalAllTime = 0;
  for (const cnt of Object.values(allTimeProjectAgg)) {
    totalAllTime += cnt;
  }
  const totalYears = Object.keys(allTimeYearAgg).length;

  // Render Helpers
  const renderProjectRow = (proj, cnt, maxVal, yearMap) => {
    const year = yearMap[proj] || '—';
    const color = yearColorMap[year] || COLORS[0];
    const percentage = Math.max(2, (cnt / maxVal) * 100);
    const showTextInside = percentage > 15; // Only show text inside if bar is wide enough

    return (
      <div key={proj} style={{ display: 'flex', alignItems: 'center', gap: 12, marginBottom: 12 }}>
        <div style={{ flex: 1, overflow: 'hidden', textOverflow: 'ellipsis', whiteSpace: 'nowrap', fontSize: 13, fontWeight: 500, color: '#334155' }} title={proj}>
          {proj}
        </div>
        <div style={{ flex: 2, position: 'relative', height: 24, background: '#f1f5f9', borderRadius: 4, overflow: 'hidden', display: 'flex', alignItems: 'center' }}>
          <div style={{ position: 'absolute', top: 0, left: 0, height: '100%', width: `${percentage}%`, background: color, transition: 'width 0.3s ease' }} />
          {/* Year inside the bar */}
          <div style={{ position: 'relative', zIndex: 1, paddingLeft: 8, fontSize: 11, fontWeight: 700, color: showTextInside ? '#ffffff' : '#64748b' }}>
            {year}
          </div>
        </div>
        <div style={{ width: 40, textAlign: 'right', fontSize: 13, fontWeight: 600, color: '#0f172a' }}>
          {cnt.toLocaleString()}
        </div>
      </div>
    );
  };



  // --- TOP SEARCHED ---
  const topData = getAggregations(topSearchFilter);
  const topProjects = Object.entries(topData.projectAgg).sort((a, b) => b[1] !== a[1] ? b[1] - a[1] : a[0].localeCompare(b[0]));
  const topMax = topProjects[0]?.[1] || 1;

  // --- LEAST SEARCHED ---
  const leastData = getAggregations(leastSearchFilter);
  const leastProjects = Object.entries(leastData.projectAgg).sort((a, b) => b[1] !== a[1] ? a[1] - b[1] : a[0].localeCompare(b[0]));
  const leastMax = leastProjects[leastProjects.length - 1]?.[1] || 1;

  // --- PIE CHART ---
  const pieData = getAggregations(pieChartFilter);
  const pieSortedYears = Object.entries(pieData.yearAgg).sort((a, b) => b[0].localeCompare(a[0]));
  const yearSlices = pieSortedYears.map(([yr, val]) => ({
    label: yr,
    value: val,
    color: yearColorMap[yr] || COLORS[0]
  }));

  if (loading && !data) {
    return <div className="an-root" style={{ padding: 40, textAlign: 'center', color: '#94a3b8' }}>Loading analytics...</div>;
  }

  const standardFilters = [
    { value: 'hour', label: 'Last Hour' },
    { value: 'day', label: 'Last 1 Day' },
    { value: 'week', label: 'Last 7 Days' },
    { value: 'month', label: 'Last 30 Days' },
    { value: 'all', label: 'All time' }
  ];

  const leastFilters = [
    { value: 'day', label: 'Last 1 Day' },
    { value: 'week', label: 'Last 7 Days' },
    { value: 'month', label: 'Last 30 Days' },
    { value: 'all', label: 'All time' }
  ];

  const mainKPIFilters = {
    hour: 'Last Hour',
    day: 'Last 24 Hours',
    week: 'Last 7 Days',
    month: 'Last 30 Days',
  };

  return (
    <div className="an-root">
      {/* ── Page Header ── */}
      <div className="an-page-header">
        <div>
          <h2 className="an-page-title">
            <Activity style={{ color: '#0078d4' }} />
            Search Analytics
          </h2>
          <div className="an-page-sub">Real-time search activity across all projects and years</div>
        </div>
        <div style={{ display: 'flex', alignItems: 'center', gap: 16 }}>
          {lastFetch && (
            <span className="an-last-updated">
              <RefreshCw size={12} /> Updated {lastFetch.toLocaleTimeString()}
            </span>
          )}
        </div>
      </div>

      {/* ── Main KPI Tabs ── */}
      <div className="an-window-tabs">
        {['hour', 'day', 'week', 'month'].map(key => (
          <button
            key={key}
            className={`an-window-tab${activeWindow === key ? ' active' : ''}`}
            onClick={() => setActiveWindow(key)}
          >
            {key === 'hour' && <Clock size={14} />}
            {key === 'day' && <Calendar size={14} />}
            {key === 'week' && <BarChart2 size={14} />}
            {key === 'month' && <TrendingUp size={14} />}
            {mainKPIFilters[key]}
          </button>
        ))}
      </div>

      {/* ── KPI Cards ── */}
      <div className="an-kpi-row">
        <KpiCard
          label={mainKPIFilters[activeWindow]}
          value={data?.windowTotals?.[activeWindow] ?? 0}
          icon={Clock}
          color="#0078d4"
          subtitle="searches in window"
        />
        <KpiCard
          label="All-Time Searches"
          value={totalAllTime.toLocaleString()}
          icon={TrendingUp}
          color="#10b981"
          subtitle="since tracking began"
        />
        <KpiCard
          label="Years Tracked"
          value={totalYears}
          icon={Layers}
          color="#8b5cf6"
          subtitle="unique project years"
        />
        <KpiCard
          label="Top Searched Year"
          value={allTimeSortedYears[0]?.[0] || '—'}
          icon={BarChart}
          color="#f59e0b"
          subtitle={allTimeSortedYears[0] ? `${allTimeSortedYears[0][1].toLocaleString()} searches` : 'no data yet'}
        />
      </div>

      {/* ── Main Dashboard Panels ── */}
      {!data || Object.keys(data.totals).length === 0 ? (
        <div className="an-panel">
          <div className="an-empty-state">
            <BarChart2 size={40} strokeWidth={1} style={{ color: '#cbd5e1' }} />
            <p>No analytics data yet.</p>
            <p style={{ fontSize: 13, color: '#94a3b8' }}>
              Perform a location-based search in Koyomail to start tracking.
            </p>
          </div>
        </div>
      ) : (
        <>
          <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20, marginBottom: 20 }}>
            {/* Top Projects Panel */}
            <div className="an-panel" style={{ display: 'flex', flexDirection: 'column' }}>
              <div className="an-panel-header">
                <span className="an-panel-title">
                  <BarChart2 size={16} style={{ color: '#0078d4' }} />
                  Top Searched Projects
                </span>
                <FilterSelect value={topSearchFilter} onChange={setTopSearchFilter} options={standardFilters} />
              </div>
              <div className="an-panel-body" style={{ flex: 1, maxHeight: 400, overflowY: 'auto', padding: '16px' }}>
                {topProjects.length === 0 && <div style={{ textAlign: 'center', color: '#94a3b8', padding: '20px 0' }}>No searches in this timeframe.</div>}
                {topProjects.slice(0, topLimit).map(([proj, cnt]) => renderProjectRow(proj, cnt, topMax, topData.projectYearMap))}
                
                {topProjects.length > topLimit && (
                  <div style={{ textAlign: 'center', marginTop: 12 }}>
                    <button 
                      className="btn btn-secondary" 
                      onClick={() => setTopLimit(l => l + 10)}
                      style={{ padding: '6px 12px', fontSize: 12 }}
                    >
                      <ChevronDown size={14} /> Load More
                    </button>
                  </div>
                )}
              </div>
            </div>

            {/* Least Searched Panel */}
            <div className="an-panel" style={{ display: 'flex', flexDirection: 'column' }}>
              <div className="an-panel-header">
                <span className="an-panel-title">
                  <BarChart2 size={16} style={{ color: '#ef4444' }} />
                  Least Searched Projects
                </span>
                <FilterSelect value={leastSearchFilter} onChange={setLeastSearchFilter} options={leastFilters} />
              </div>
              <div className="an-panel-body" style={{ flex: 1, maxHeight: 400, overflowY: 'auto', padding: '16px' }}>
                {leastProjects.length === 0 && <div style={{ textAlign: 'center', color: '#94a3b8', padding: '20px 0' }}>No searches in this timeframe.</div>}
                {leastProjects.slice(0, leastLimit).map(([proj, cnt]) => renderProjectRow(proj, cnt, leastMax, leastData.projectYearMap))}
                
                {leastProjects.length > leastLimit && (
                  <div style={{ textAlign: 'center', marginTop: 12 }}>
                    <button 
                      className="btn btn-secondary" 
                      onClick={() => setLeastLimit(l => l + 10)}
                      style={{ padding: '6px 12px', fontSize: 12 }}
                    >
                      <ChevronDown size={14} /> Load More
                    </button>
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Searches by Year Panel (Pie Chart) */}
          <div className="an-panel" style={{ maxWidth: 600, margin: '0 auto', marginBottom: 20 }}>
            <div className="an-panel-header">
              <span className="an-panel-title">
                <Layers size={16} style={{ color: '#8b5cf6' }} />
                Searches by Year
              </span>
              <FilterSelect value={pieChartFilter} onChange={setPieChartFilter} options={standardFilters} />
            </div>
            <div className="an-panel-body" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', minHeight: 300, padding: '20px' }}>
              {yearSlices.length === 0 ? (
                <div style={{ color: '#94a3b8' }}>No searches in this timeframe.</div>
              ) : (
                <>
                  <DonutChart slices={yearSlices} size={200} />
                  <div style={{ marginTop: 24, display: 'flex', flexWrap: 'wrap', gap: 16, justifyContent: 'center' }}>
                    {yearSlices.map(slice => (
                      <div key={slice.label} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                        <div style={{ width: 12, height: 12, borderRadius: '50%', background: slice.color }}></div>
                        <span style={{ fontSize: 13, fontWeight: 600, color: '#334155' }}>{slice.label}</span>
                        <span style={{ fontSize: 12, color: '#64748b' }}>({slice.value.toLocaleString()})</span>
                      </div>
                    ))}
                  </div>
                </>
              )}
            </div>
          </div>

          {/* Year @ Project Breakdown */}
          <div style={{ marginTop: 20 }}>
            <h3 style={{ fontSize: 16, fontWeight: 700, color: '#1e293b', marginBottom: 16 }}>
              Breakdown by Year (All Time)
            </h3>
            <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: 16 }}>
              {allTimeSortedYears.map(([year]) => {
                const yearProjs = Object.entries(data.totals[year]).sort((a,b) => b[1] !== a[1] ? b[1]-a[1] : a[0].localeCompare(b[0]));
                const yrMax = yearProjs[0]?.[1] || 1;
                const yearColor = yearColorMap[year];
                const visibleProjs = yearProjs.slice(0, 15);
                const extraProjs = yearProjs.length - 15;

                return (
                  <div key={year} className="an-panel" style={{ borderTop: `4px solid ${yearColor}` }}>
                    <div className="an-panel-header" style={{ padding: '12px 16px' }}>
                      <span className="an-panel-title" style={{ fontSize: 14 }}>{year}</span>
                      <span className="an-panel-sub">{allTimeYearAgg[year].toLocaleString()} searches</span>
                    </div>
                    <div className="an-panel-body" style={{ padding: '12px 16px', maxHeight: 280, overflowY: 'auto' }}>
                      {visibleProjs.map(([proj, cnt]) => (
                        <div key={proj} className="an-project-row" style={{ padding: '6px 0', borderBottom: '1px solid #f1f5f9' }}>
                          <div className="an-project-name" title={proj} style={{ fontSize: 12 }}>{proj}</div>
                          <div className="an-project-bar-wrap" style={{ height: 6 }}>
                            <div className="an-project-bar-fill" style={{ width: `${Math.round((cnt/yrMax)*100)}%`, background: yearColor }} />
                          </div>
                          <div className="an-project-count" style={{ fontSize: 12 }}>{cnt}</div>
                        </div>
                      ))}
                      {extraProjs > 0 && (
                        <div style={{ padding: '12px 0', textAlign: 'center', fontSize: 12, color: '#94a3b8' }}>
                          +{extraProjs.toLocaleString()} more projects in {year}
                        </div>
                      )}
                    </div>
                  </div>
                );
              })}
            </div>
          </div>
        </>
      )}
    </div>
  );
}
