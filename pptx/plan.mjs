// Click states of one slide → shapes plus the animation effects that move between the states.
// A shape that only changes colour or weight stays one shape with emphasis effects; a shape
// whose content changes becomes stacked copies that appear and disappear.
import { cmpKey } from './scene.mjs'

const round1 = v => Math.round(v * 10) / 10
function canon(o, drop = ['order', 'name']) {
  if (typeof o === 'number') return round1(o)
  if (Array.isArray(o)) return o.map(v => canon(v, drop))
  if (o && typeof o === 'object') {
    const out = {}
    for (const k of Object.keys(o).sort()) if (!drop.includes(k) && o[k] !== undefined) out[k] = canon(o[k], drop)
    return out
  }
  return o
}
const sig = o => JSON.stringify(canon(o))

// The spec with everything an emphasis effect can change removed.
function strip(o) {
  o = structuredClone(o)
  delete o.order
  if (o.type === 'shape') {
    const f = o.fill
    o.fill = f && !f.pattern && !f.gradient && (f.alpha ?? 1) < 0.999 ? { alpha: f.alpha } : (f && (f.pattern || f.gradient) ? f : null)
    o.line = null
    for (const p of o.text?.paras || []) {
      for (const r of p.runs) { delete r.color; delete r.b }
    }
  } else if (o.type === 'cxn') o.line = { ...o.line, color: null }
  return sig(o)
}

function minor(c1, c2) {
  if (!c1 || !c2) return c1 === c2
  return Math.max(...[0, 2, 4].map(i => Math.abs(parseInt(c1.slice(i, i + 2), 16) - parseInt(c2.slice(i, i + 2), 16)))) <= 10
}

// Per paragraph, the value of attr in each state ('base' when unchanged); null when one effect cannot express it.
function textSequences(specs, attr) {
  const base = specs[0]
  if (base.type !== 'shape' || !base.text) return []
  const norm = v => attr === 'b' ? !!v : v
  const out = []
  for (let j = 0; j < base.text.paras.length; j++) {
    const bvals = base.text.paras[j].runs.filter(r => r.t !== undefined).map(r => norm(r[attr]))
    const seq = []
    for (const s of specs) {
      const vals = s.text.paras[j].runs.filter(r => r.t !== undefined).map(r => norm(r[attr]))
      const same = JSON.stringify(vals) === JSON.stringify(bvals) || (attr === 'color' && vals.every((v, k) => minor(v, bvals[k])))
      if (same) seq.push('base')
      else if (new Set(vals).size === 1) seq.push(vals[0])
      else return null
    }
    const firstChange = seq.findIndex(v => v !== 'base')
    if (firstChange >= 0 && seq.slice(firstChange).includes('base') && new Set(bvals).size > 1) return null // back to a mixed base
    out.push([seq, bvals.length ? bvals[0] : null])
  }
  return out
}

// Effects for click k: one for the whole shape if every paragraph moves to one value, else per paragraph.
function textEffects(seqs, k, kind) {
  const changes = []
  seqs.forEach(([seq, b0], j) => { if (seq[k] !== seq[k - 1]) changes.push([j, seq[k] !== 'base' ? seq[k] : b0]) })
  if (!changes.length) return []
  const texty = seqs.map(([, b0], j) => b0 !== null ? j : -1).filter(j => j >= 0)
  if (new Set(changes.map(c => JSON.stringify(c[1]))).size === 1 && JSON.stringify(changes.map(c => c[0])) === JSON.stringify(texty)) return [[kind, changes[0][1]]]
  return changes.map(([j, v]) => [kind, v, j])
}

const lineCore = l => l ? JSON.stringify([round1(l.w), l.dash || null, Math.round((l.alpha ?? 1) * 100) / 100]) : null
const GEOM = ['x', 'y', 'w', 'h', 'x1', 'y1', 'x2', 'y2']

// cur with every number within a pixel or two of prev replaced by prev's value.
function settleValue(cur, prev, key = '') {
  if (Array.isArray(cur) && Array.isArray(prev) && cur.length === prev.length) return cur.map((v, i) => settleValue(v, prev[i], key))
  if (cur && prev && typeof cur === 'object' && typeof prev === 'object' && !Array.isArray(cur)) {
    // outline widths and colours are what a highlight changes; never smooth those
    const out = {}
    for (const [k, v] of Object.entries(cur)) out[k] = k in prev && !['line', 'borders', 'fill', 'bullet', 'shadow'].includes(k) ? settleValue(v, prev[k], k) : v
    return out
  }
  if (typeof cur === 'number' && typeof prev === 'number') {
    if (['sz', 'alpha', 'adj', 'start'].includes(key)) return cur
    const tol = GEOM.includes(key) ? 3 : Math.abs(cur) > 1000 ? 2000 : 1.5
    return Math.abs(cur - prev) <= tol ? prev : cur
  }
  return cur
}
function settle(specs) {
  const out = [...specs]
  let prev = null
  out.forEach((s, i) => {
    if (!s) { prev = null; return }
    if (prev && prev.type === s.type) out[i] = s = settleValue(s, prev)
    prev = s
  })
  return out
}

function planVersions(specs, N) {
  const out = []
  for (let i = 0; i < specs.length;) {
    if (!specs[i]) { i++; continue }
    let j = i
    const s = sig(specs[i])
    while (j + 1 < specs.length && specs[j + 1] && sig(specs[j + 1]) === s) j++
    out.push({ spec: specs[i], appear: i > 0 ? i : null, disappear: j + 1 <= N ? j + 1 : null, effects: {} })
    i = j + 1
  }
  return out
}

// One object key across all click states → the shapes to emit.
export function planKey(specs, N, bg) {
  specs = settle(specs)
  const present = specs.map((s, i) => s ? i : -1).filter(i => i >= 0)
  const a = present[0], b = present[present.length - 1]
  const contiguous = present.length === b - a + 1
  const appear = a > 0 ? a : null, disappear = b + 1 <= N ? b + 1 : null
  const distinct = new Set(specs.filter(Boolean).map(sig))
  if (contiguous && distinct.size === 1) return [{ spec: specs[a], appear, disappear, effects: {} }]
  let base = specs[a]
  const translucent = specs.some(s => s && (s.fill?.alpha ?? 1) < 0.999)
  const range = specs.slice(a, b + 1)
  if (contiguous && !translucent && ['shape', 'cxn'].includes(base.type) && new Set(range.map(strip)).size === 1) {
    const fontSeq = textSequences(range, 'color'), boldSeq = textSequences(range, 'b')
    if (fontSeq && boldSeq) {
      const lines = range.map(s => s.line)
      const overlay = new Set(lines.map(lineCore)).size > 1
      const effects = {}
      for (let k = a + 1; k <= b; k++) {
        const prev = specs[k - 1], cur = specs[k], ev = []
        if (cur.type === 'shape') {
          const pf = prev.fill?.color, cf = cur.fill?.color
          if (pf !== cf) ev.push(['fill', cf || bg])
        }
        const pl = prev.line?.color, cl = cur.line?.color
        if (!overlay && pl !== cl && cl) ev.push(['line', cl])
        ev.push(...textEffects(fontSeq, k - a, 'font'), ...textEffects(boldSeq, k - a, 'bold'))
        if (ev.length) effects[k] = ev
      }
      if (overlay && base.type !== 'shape') return planVersions(specs, N)
      // the shape keeps its lightest outline; heavier outlines are drawn over it while they apply
      if (overlay) base = { ...base, line: lines.reduce((m, l) => (l ? l.w : 0) < (m ? m.w : 0) ? l : m) }
      const out = [{ spec: base, appear, disappear, effects }]
      if (overlay) {
        const bl = lineCore(base.line)
        const ospecs = specs.map((s, i) => {
          if (i < a || i > b) return null
          const l = s.line
          if (lineCore(l) === bl && !(l && base.line && l.color !== base.line.color)) return null
          const { text, fill, ...rest } = s
          return { ...rest, fill: null, text: null, shadow: null, name: 'Highlight Outline', order: [...s.order, 9999] }
        })
        if (ospecs.some(Boolean)) out.push(...planVersions(ospecs, N))
      }
      return out
    }
  }
  return planVersions(specs, N)
}

// The object as it looks after click k once its effects have played (null when hidden).
export function simulate({ spec, appear, disappear, effects }, k) {
  if ((appear != null && k < appear) || (disappear != null && k >= disappear)) return null
  spec = structuredClone(spec)
  for (const c of Object.keys(effects).map(Number).sort((x, y) => x - y)) {
    if (c > k) continue
    for (const [kind, value, para] of effects[c]) {
      let paras = spec.text?.paras || []
      if (para !== undefined) paras = [paras[para]]
      const runs = paras.flatMap(p => p.runs.filter(r => r.t !== undefined))
      if (kind === 'fill') spec.fill = { color: value, alpha: 1 }
      else if (kind === 'line' && spec.line) spec.line = { ...spec.line, color: value }
      else if (kind === 'font') for (const r of runs) r.color = value
      else if (kind === 'bold') for (const r of runs) r.b = value
    }
  }
  return spec
}

// All keys of a slide → ordered emission plan.
export function planSlide(perState, bg) {
  const N = perState.length - 1
  const keys = []
  const seen = new Set()
  for (const so of perState) for (const k of so.keys()) if (!seen.has(k)) { seen.add(k); keys.push(k) }
  const plan = [], copies = []
  for (const key of keys) {
    const specs = perState.map(so => so.get(key) || null)
    const planned = planKey(specs, N, bg)
    const texty = planned.filter(p => ['shape', 'ph', 'table'].includes(p.spec.type) && (p.spec.text || p.spec.paras || p.spec.type === 'table'))
    if (texty.length > 1) copies.push(key)
    for (const p of planned) plan.push({ key, ...p })
  }
  plan.sort((p, q) => cmpKey(p.spec.order, q.spec.order) || (p.appear ?? -1) - (q.appear ?? -1))
  return { plan, copies, N }
}
