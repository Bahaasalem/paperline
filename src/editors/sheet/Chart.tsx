import type { ChartType } from './model';
import { CHART_COLORS, type ChartData } from './features';

const FONT = "Inter, 'IBM Plex Sans Arabic', sans-serif";

function niceScale(lo: number, hi: number, ticks = 5) {
  if (lo === hi) { hi = lo + 1; lo = Math.min(0, lo); }
  const span = hi - lo;
  const raw = span / ticks;
  const mag = 10 ** Math.floor(Math.log10(raw));
  const step = [1, 2, 2.5, 5, 10].map((m) => m * mag).find((s) => span / s <= ticks) ?? mag * 10;
  const min = Math.floor(lo / step) * step;
  const max = Math.ceil(hi / step) * step;
  const out: number[] = [];
  for (let v = min; v <= max + step / 2; v += step) out.push(+v.toPrecision(10));
  return out;
}

const short = (n: number) => {
  const a = Math.abs(n);
  if (a >= 1e9) return `${+(n / 1e9).toFixed(1)}B`;
  if (a >= 1e6) return `${+(n / 1e6).toFixed(1)}M`;
  if (a >= 1e4) return `${+(n / 1e3).toFixed(1)}K`;
  return String(+n.toPrecision(6));
};

/** An SVG chart drawn from sheet data. Used on the grid, in print and in PDF export. */
export function ChartSvg({ type, data, w, h, title, legend = true }: { type: ChartType; data: ChartData; w: number; h: number; title?: string; legend?: boolean }) {
  const titleH = title ? 30 : 10;
  const pie = type === 'pie' || type === 'doughnut';
  const legendItems = pie ? data.cats : data.series.map((s) => s.name);
  const legendH = legend && legendItems.length > (pie ? 0 : 1) ? 26 : 0;
  const body = { x: 0, y: titleH, w, h: h - titleH - legendH };

  let plot: React.ReactNode = null;
  if (pie) {
    const vals = (data.series[0]?.values ?? []).map((v) => Math.max(0, v ?? 0));
    const total = vals.reduce((s, v) => s + v, 0) || 1;
    const cx = w / 2, cy = body.y + body.h / 2, R = Math.max(10, Math.min(body.w, body.h) / 2 - 8);
    let a0 = -Math.PI / 2;
    plot = (
      <g>
        {vals.map((v, i) => {
          const a1 = a0 + (v / total) * Math.PI * 2;
          const large = a1 - a0 > Math.PI ? 1 : 0;
          const p = (a: number, r: number) => `${cx + r * Math.cos(a)} ${cy + r * Math.sin(a)}`;
          const mid = (a0 + a1) / 2;
          const d = v >= total - 1e-9
            ? `M ${cx - R} ${cy} a ${R} ${R} 0 1 0 ${2 * R} 0 a ${R} ${R} 0 1 0 ${-2 * R} 0`
            : `M ${cx} ${cy} L ${p(a0, R)} A ${R} ${R} 0 ${large} 1 ${p(a1, R)} Z`;
          const label = v / total >= 0.06 ? `${Math.round((v / total) * 100)}%` : '';
          const lp = p(mid, type === 'doughnut' ? R * 0.78 : R * 0.62).split(' ').map(Number);
          a0 = a1;
          return (
            <g key={i}>
              <path d={d} fill={CHART_COLORS[i % CHART_COLORS.length]} stroke="#fff" strokeWidth={1.5} />
              {label && <text x={lp[0]} y={lp[1]} fill="#fff" fontSize={11} fontWeight={600} textAnchor="middle" dominantBaseline="central">{label}</text>}
            </g>
          );
        })}
        {type === 'doughnut' && <circle cx={cx} cy={cy} r={R * 0.55} fill="#fff" />}
      </g>
    );
  } else {
    const all = data.series.flatMap((s) => s.values.filter((v): v is number => v !== null));
    const scatter = type === 'scatter';
    const ticks = niceScale(Math.min(0, ...all), Math.max(0, ...all));
    const lo = ticks[0], hi = ticks[ticks.length - 1];
    const horiz = type === 'bar';
    const axisW = Math.max(...ticks.map((t) => short(t).length)) * 6.5 + 12;
    const catH = 22;
    const area = horiz
      ? { x: Math.min(110, Math.max(...data.cats.map((c) => c.length), 1) * 6.5 + 14), y: body.y + 4, w: 0, h: body.h - catH - 4 }
      : { x: axisW, y: body.y + 4, w: 0, h: body.h - catH - 4 };
    area.w = w - area.x - 14;
    const scale = (v: number) => (v - lo) / (hi - lo || 1);
    const n = data.cats.length || 1;
    const ns = data.series.length || 1;
    const band = (horiz ? area.h : area.w) / n;
    const grid = ticks.map((t) => {
      const pos = horiz ? area.x + scale(t) * area.w : area.y + area.h - scale(t) * area.h;
      return horiz ? (
        <g key={t}>
          <line x1={pos} x2={pos} y1={area.y} y2={area.y + area.h} stroke="#e5e7eb" />
          <text x={pos} y={area.y + area.h + 15} fontSize={10.5} fill="#6b7280" textAnchor="middle">{short(t)}</text>
        </g>
      ) : (
        <g key={t}>
          <line x1={area.x} x2={area.x + area.w} y1={pos} y2={pos} stroke="#e5e7eb" />
          <text x={area.x - 6} y={pos} fontSize={10.5} fill="#6b7280" textAnchor="end" dominantBaseline="central">{short(t)}</text>
        </g>
      );
    });
    const every = Math.max(1, Math.ceil(n / Math.max(1, Math.floor((horiz ? area.h : area.w) / 46))));
    const cats = data.cats.map((c, i) =>
      i % every ? null : horiz ? (
        <text key={i} x={area.x - 6} y={area.y + band * (i + 0.5)} fontSize={10.5} fill="#374151" textAnchor="end" dominantBaseline="central">{c.slice(0, 16)}</text>
      ) : (
        <text key={i} x={area.x + band * (i + 0.5)} y={area.y + area.h + 15} fontSize={10.5} fill="#374151" textAnchor="middle">{c.slice(0, 14)}</text>
      ),
    );
    const zero = horiz ? area.x + scale(Math.max(lo, 0)) * area.w : area.y + area.h - scale(Math.max(lo, 0)) * area.h;
    let marks: React.ReactNode;
    if (type === 'column' || type === 'bar') {
      const gap = band * 0.2;
      const bw = (band - gap) / ns;
      marks = data.series.map((s, si) =>
        s.values.map((v, i) => {
          if (v === null) return null;
          const p = horiz ? area.x + scale(v) * area.w : area.y + area.h - scale(v) * area.h;
          const o = band * i + gap / 2 + bw * si;
          return horiz ? (
            <rect key={`${si}-${i}`} x={Math.min(p, zero)} y={area.y + o} width={Math.abs(p - zero)} height={Math.max(1, bw - 1)} rx={2} fill={CHART_COLORS[si % CHART_COLORS.length]} />
          ) : (
            <rect key={`${si}-${i}`} x={area.x + o} y={Math.min(p, zero)} width={Math.max(1, bw - 1)} height={Math.abs(p - zero)} rx={2} fill={CHART_COLORS[si % CHART_COLORS.length]} />
          );
        }),
      );
    } else {
      marks = data.series.map((s, si) => {
        const color = CHART_COLORS[si % CHART_COLORS.length];
        const pts = s.values.map((v, i) => (v === null ? null : [area.x + band * (i + 0.5), area.y + area.h - scale(v) * area.h] as const));
        const line = pts.filter(Boolean).map((p, i) => `${i ? 'L' : 'M'} ${p![0]} ${p![1]}`).join(' ');
        const firstPt = pts.find(Boolean), lastPt = [...pts].reverse().find(Boolean);
        return (
          <g key={si}>
            {type === 'area' && firstPt && lastPt && <path d={`${line} L ${lastPt[0]} ${zero} L ${firstPt[0]} ${zero} Z`} fill={color} opacity={0.22} />}
            {!scatter && <path d={line} fill="none" stroke={color} strokeWidth={2.4} strokeLinejoin="round" strokeLinecap="round" />}
            {(scatter || n <= 24) && pts.map((p, i) => p && <circle key={i} cx={p[0]} cy={p[1]} r={scatter ? 4 : 3} fill={scatter ? color : '#fff'} stroke={color} strokeWidth={2} />)}
          </g>
        );
      });
    }
    plot = (
      <g>
        {grid}
        {horiz
          ? <line x1={zero} x2={zero} y1={area.y} y2={area.y + area.h} stroke="#9ca3af" />
          : <line x1={area.x} x2={area.x + area.w} y1={zero} y2={zero} stroke="#9ca3af" />}
        {marks}
        {cats}
      </g>
    );
  }

  let lx = 0;
  const items = legendH
    ? legendItems.map((name, i) => {
        const x = lx;
        lx += Math.min(140, name.length * 6.6 + 26);
        return (
          <g key={i} transform={`translate(${x} 0)`}>
            <rect width={10} height={10} y={-5} rx={2} fill={CHART_COLORS[i % CHART_COLORS.length]} />
            <text x={15} fontSize={11} fill="#374151" dominantBaseline="central">{name.slice(0, 18)}</text>
          </g>
        );
      })
    : null;

  return (
    <svg width={w} height={h} viewBox={`0 0 ${w} ${h}`} className="chart-svg" style={{ fontFamily: FONT, direction: 'ltr' }}>
      <rect width={w} height={h} fill="#fff" />
      {title && <text x={w / 2} y={20} fontSize={14} fontWeight={600} fill="#111827" textAnchor="middle">{title}</text>}
      {plot}
      {items && <g transform={`translate(${Math.max(8, (w - lx) / 2)} ${h - 13})`}>{items}</g>}
    </svg>
  );
}
