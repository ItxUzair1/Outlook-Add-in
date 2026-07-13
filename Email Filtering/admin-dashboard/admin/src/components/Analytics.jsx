import { useState, useEffect, useRef, useCallback } from 'react';
import {
  BarChart2, TrendingUp, Clock, Calendar, BarChart,
  RefreshCw, Layers, Activity
} from 'lucide-react';

const API_BASE_URL = import.meta.env.VITE_API_BASE_URL || 'http://localhost:4001/api';
const POLL_MS = 2500;

// ─── tiny pie / donut chart ──────────────────────────────────────────────────
function DonutChart({ slices, size = 100 }) {
  if (!slices || slices.length === 0) return null;
  const total = slices.reduce((s, sl) => s + sl.value, 0);
  if (total === 0) return null;

  const COLORS = [
    '#0078d4', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#06b6d4', '#ec4899', '#84cc16',
    '#6366f1', '#14b8a6', '#f97316', '#d946ef', '#0ea5e9', '#8b5cf6', '#f43f5e', '#a855f7',
    '#22c55e', '#eab308', '#3b82f6', '#10b981', '#f43f5e', '#64748b'
  ];
  const cx = size / 2, cy = size / 2, r = size / 2 - 6, ri = r * 0.55;
  // Pre-compute start angles with reduce so no variable is mutated inside map
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
    
    // Percentage text
    const pct = Math.round((sl.value / total) * 100);
    const midAngle = startAngle + sweep / 2;
    const textR = r * 0.77; // Position text exactly midway in the donut thickness
    const tx = cx + textR * Math.cos(midAngle);
    const ty = cy + textR * Math.sin(midAngle);

    return (
      <g key={i}>
        <path d={d} fill={COLORS[i % COLORS.length]}>
          <title>{`${sl.label}: ${sl.value}`}</title>
        </path>
        {pct >= 5 && ( // Only show percent if slice is big enough
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


// ─── Analytics Page ────────────────────────────────────────────────────────────
export default function Analytics() {
  const [data, setData] = useState(null);
  const [loading, setLoading] = useState(true);
  const [activeWindow, setActiveWindow] = useState('day'); // hour, day, week, month
  const [lastFetch, setLastFetch] = useState(null);
  const intervalRef = useRef(null);

  const fetchAnalytics = useCallback(async () => {
    try {
      const resp = await fetch(`${API_BASE_URL}/analytics`);
      if (resp.ok) {
        const json = await resp.json();
        setData(json);
        setLastFetch(new Date());
      }
    } catch {
      // silently fail — indexer might not be running
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

  const windowLabels = {
    hour: 'Last Hour',
    day: 'Last 24 Hours',
    week: 'Last 7 Days',
    month: 'Last 30 Days',
  };

  let totalAllTime = 0;
  let totalYears = 0;
  const projectAgg = {};
  const yearAgg = {};
  const projectYearMap = {}; // Map project to the year it had the most searches in

  if (data?.totals) {
    totalYears = Object.keys(data.totals).length;
    for (const [year, projects] of Object.entries(data.totals)) {
      let yearSum = 0;
      for (const [proj, cnt] of Object.entries(projects)) {
        totalAllTime += cnt;
        yearSum += cnt;
        projectAgg[proj] = (projectAgg[proj] || 0) + cnt;
        
        // Keep track of the year where the project has the highest count
        if (!projectYearMap[proj] || cnt > (data.totals[projectYearMap[proj]][proj] || 0)) {
            projectYearMap[proj] = year;
        }
      }
      yearAgg[year] = yearSum;
    }
  }

  // Sort projects and years - use localeCompare for stable sorting when counts are equal
  const topProjects = Object.entries(projectAgg).sort((a, b) => b[1] !== a[1] ? b[1] - a[1] : a[0].localeCompare(b[0]));
  const sortedYears = Object.entries(yearAgg).sort((a, b) => b[0].localeCompare(a[0]));
  
  const COLORS = [
    '#0078d4', '#10b981', '#f59e0b', '#8b5cf6', '#ef4444', '#06b6d4', '#ec4899', '#84cc16',
    '#6366f1', '#14b8a6', '#f97316', '#d946ef', '#0ea5e9', '#8b5cf6', '#f43f5e', '#a855f7',
    '#22c55e', '#eab308', '#3b82f6', '#10b981', '#f43f5e', '#64748b'
  ];
  const yearSlices = sortedYears.map(([yr, val], i) => ({
    label: yr,
    value: val,
    color: COLORS[i % COLORS.length]
  }));
  
  const yearColorMap = {};
  yearSlices.forEach(slice => {
      yearColorMap[slice.label] = slice.color;
  });


  const maxProjVal = topProjects[0]?.[1] || 1;

  if (loading && !data) {
    return <div className="an-root" style={{ padding: 40, textAlign: 'center', color: '#94a3b8' }}>Loading analytics...</div>;
  }

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

      {/* ── Tabs ── */}
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
            {windowLabels[key]}
          </button>
        ))}
      </div>

      {/* ── KPI Cards ── */}
      <div className="an-kpi-row">
        <KpiCard
          label={windowLabels[activeWindow]}
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
          value={sortedYears[0]?.[0] || '—'}
          icon={BarChart}
          color="#f59e0b"
          subtitle={sortedYears[0] ? `${sortedYears[0][1].toLocaleString()} searches` : 'no data yet'}
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
        <div style={{ display: 'grid', gridTemplateColumns: '1fr 1fr', gap: 20 }}>
          
          {/* Top Projects Panel */}
          <div className="an-panel">
            <div className="an-panel-header">
              <span className="an-panel-title">
                <BarChart2 size={16} style={{ color: '#0078d4' }} />
                Top Searched Projects
              </span>
              <span className="an-panel-sub">All time</span>
            </div>
            <div className="an-panel-body" style={{ maxHeight: 400, overflowY: 'auto' }}>
              {topProjects.map(([proj, cnt]) => (
                <div key={proj} className="an-project-row">
                  <div className="an-project-name" title={proj}>{proj}</div>
                  <div className="an-project-bar-wrap">
                    <div
                      className="an-project-bar-fill"
                      style={{
                        width: `${Math.round((cnt / maxProjVal) * 100)}%`,
                        background: yearColorMap[projectYearMap[proj]] || COLORS[0]
                      }}
                    />
                  </div>
                  <div className="an-project-count">{cnt.toLocaleString()}</div>
                </div>
              ))}
            </div>
          </div>

          {/* Searches by Year Panel */}
          <div className="an-panel">
            <div className="an-panel-header">
              <span className="an-panel-title">
                <Layers size={16} style={{ color: '#8b5cf6' }} />
                Searches by Year
              </span>
              <span className="an-panel-sub">Distribution of searches</span>
            </div>
            <div className="an-panel-body" style={{ display: 'flex', flexDirection: 'column', alignItems: 'center', justifyContent: 'center', height: '100%', minHeight: 300 }}>
              <DonutChart slices={yearSlices} size={200} />
              
              <div style={{ marginTop: 24, display: 'flex', flexWrap: 'wrap', gap: 16, justifyContent: 'center' }}>
                {yearSlices.map(slice => (
                  <div key={slice.label} style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
                    <div style={{ width: 12, height: 12, borderRadius: '50%', background: slice.color }}></div>
                    <span style={{ fontSize: 13, fontWeight: 600, color: '#334155' }}>{slice.label}</span>
                    <span style={{ fontSize: 12, color: '#64748b' }}>({slice.value})</span>
                  </div>
                ))}
              </div>
            </div>
          </div>

        </div>

        {/* Year @ Project Breakdown */}
        <div style={{ marginTop: 10 }}>
          <h3 style={{ fontSize: 16, fontWeight: 700, color: '#1e293b', marginBottom: 16, marginTop: 10 }}>
            Breakdown by Year
          </h3>
          <div style={{ display: 'grid', gridTemplateColumns: 'repeat(auto-fit, minmax(320px, 1fr))', gap: 16 }}>
            {sortedYears.map(([year]) => {
              const yearProjs = Object.entries(data.totals[year]).sort((a,b) => b[1] !== a[1] ? b[1]-a[1] : a[0].localeCompare(b[0]));
              const yrMax = yearProjs[0]?.[1] || 1;
              const yearColor = yearColorMap[year];
              const visibleProjs = yearProjs.slice(0, 15);
              const extraProjs = yearProjs.length - 15;

              return (
                <div key={year} className="an-panel" style={{ borderTop: `4px solid ${yearColor}` }}>
                  <div className="an-panel-header" style={{ padding: '12px 16px' }}>
                    <span className="an-panel-title" style={{ fontSize: 14 }}>{year}</span>
                    <span className="an-panel-sub">{yearAgg[year].toLocaleString()} searches</span>
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
