// Extracted DOM items of one click state → PowerPoint object specs, in slide pixels.
import { parsePath, transformSubs, pointsOf, bbox, simplifyPolyline, apply } from './geom.mjs'

// Single line spacing of each font relative to its size, as PowerPoint lays it out.
const LINE_FACTOR = { Arial: 1.149, Georgia: 1.136, Consolas: 1.171, 'Cambria Math': 1.172, 'Times New Roman': 1.15, 'Courier New': 1.133, Calibri: 1.22, Cambria: 1.172, Verdana: 1.215, Tahoma: 1.207, 'Trebuchet MS': 1.16, 'Segoe UI': 1.33 }
const lineFactor = font => LINE_FACTOR[font] ?? 1.15
const SLACK = 3

export const hexc = c => c ? [c.r, c.g, c.b].map(v => Math.max(0, Math.min(255, Math.round(v))).toString(16).padStart(2, '0')).join('').toUpperCase() : null
const r1 = v => Math.round(v * 10) / 10
const r2 = v => Math.round(v * 100) / 100

// DOM pre-order sort key from an extracted id path.
export function pathKey(path) {
  const [main, suffix] = path.split('#')
  const out = main.split('/').map(c => { const m = c.match(/(\d+)$/); return m ? parseInt(m[1]) : 0 })
  if (suffix !== undefined) out.push(parseInt(suffix) || 0)
  return out
}
export function cmpKey(a, b) {
  for (let i = 0; i < Math.min(a.length, b.length); i++) if (a[i] !== b[i]) return a[i] - b[i]
  return a.length - b.length
}
const byId = (p, q) => cmpKey(pathKey(p.id), pathKey(q.id))

// Font sizes land on half points, as typed in PowerPoint; round down so lines never get longer.
export const snap = px => Math.max(Math.floor(px * 0.75 * 2 + 0.2) / 2, 1) / 0.75

// ---------------------------------------------------------------- text ----

export function normRuns(p) {
  const pre = p.pre || (p.whiteSpace || '').startsWith('pre')
  let runs = []
  for (const r of p.runs) {
    if (r.br) { runs.push({ br: true }); continue }
    if (r.math !== undefined) { runs.push({ omml: r.omml || '', tex: r.tex || '', mathId: r.mathId, sz: snap(r.size), color: hexc(r.color), display: r.display }); continue }
    let t = r.t
    const ws = r.ws || ''
    if (!(pre || ws.startsWith('pre') || ws === 'break-spaces')) t = t.replace(/[ \t\n\r\f]+/g, ' ')
    else if (ws === 'pre-line') t = t.replace(/[ \t]+/g, ' ')
    if (t === '') continue
    runs.push({
      t, sz: snap(r.size), b: r.weight >= 600, i: r.italic, u: r.u, s: r.s, color: hexc(r.color), font: r.font, va: r.va,
      hl: r.hl ? hexc(r.hl) : null, spc: r.ls ? r2(r.ls) : 0, caps: r.caps, link: r.link,
    })
  }
  if (!pre) {
    // collapse spaces across run boundaries and trim paragraph ends / around breaks
    let prevSpace = true
    for (const r of runs) {
      if (r.br) { prevSpace = true; continue }
      if (r.omml !== undefined) { prevSpace = false; continue }
      if (prevSpace) r.t = r.t.replace(/^ +/, '')
      if (r.t) prevSpace = r.t.endsWith(' ')
    }
    for (let i = runs.length - 1; i >= 0; i--) {
      const r = runs[i]
      if (r.br || r.omml !== undefined) break
      r.t = r.t.replace(/ +$/, '')
      if (r.t) break
    }
    runs.forEach((r, i) => { if (r.br && i > 0 && runs[i - 1].t !== undefined) runs[i - 1].t = runs[i - 1].t.replace(/ +$/, '') })
    runs = runs.filter(r => r.br || r.omml !== undefined || r.t)
    while (runs.length && runs[runs.length - 1].br) runs.pop()
  }
  // merge adjacent runs with identical formatting
  const same = (a, b) => JSON.stringify({ ...a, t: 0 }) === JSON.stringify({ ...b, t: 0 })
  const merged = []
  for (const r of runs) {
    const last = merged[merged.length - 1]
    if (last && r.t !== undefined && last.t !== undefined && same(r, last)) last.t += r.t
    else merged.push({ ...r })
  }
  return merged
}

function mainFont(p, runs) {
  const tally = new Map()
  for (const r of runs) if (r.t !== undefined && !r.va) {
    const k = r.font + '|' + r2(r.sz)
    tally.set(k, (tally.get(k) || 0) + r.t.length)
  }
  if (tally.size) {
    const [k] = [...tally.entries()].reduce((a, b) => b[1] > a[1] ? b : a)
    const [font, sz] = k.split('|')
    return [font, runs.find(r => r.font === font && r2(r.sz) === parseFloat(sz)).sz]
  }
  return [runs.find(r => r.font)?.font || 'Arial', snap(p.fs)]
}

// Superscripts sit higher than their line; fold overlapping boxes into one line.
function mergeLines(lines) {
  const out = []
  for (const l of [...lines].sort((a, b) => a.top - b.top)) {
    const m = out.find(m => Math.min(m.bottom, l.bottom) - Math.max(m.top, l.top) > 0.5 * Math.min(m.bottom - m.top, l.bottom - l.top))
    if (m) { m.top = Math.min(m.top, l.top); m.bottom = Math.max(m.bottom, l.bottom); m.left = Math.min(m.left, l.left); m.right = Math.max(m.right, l.right); m.maxh = Math.max(m.maxh, l.maxh) }
    else out.push({ ...l })
  }
  return out
}

function paraGeom(p) {
  p.lines = mergeLines(p.lines)
  const runs = p._runs
  const [font, fs] = mainFont(p, runs)
  const sized = runs.filter(r => r.t !== undefined && !r.va && r.t.trim())
  const fmax = sized.length ? Math.max(...sized.map(r => r.sz)) : fs
  const fmaxFont = sized.find(r => r.sz === fmax)?.font || font
  const lh = p.lh || fs * lineFactor(font)
  const lines = p.lines, c = p.content, n = Math.max(1, lines.length)
  let top, bottom
  if (p.onlyChild && Math.abs(c.h - n * lh) < lh * 0.45) { top = c.y; bottom = c.y + c.h }
  else if (lines.length) {
    const c0 = (lines[0].top + lines[0].bottom) / 2, c1 = (lines[lines.length - 1].top + lines[lines.length - 1].bottom) / 2
    top = c0 - lh / 2; bottom = c1 + lh / 2
  } else { top = c.y; bottom = c.y + lh }
  const tl = lines.length ? Math.min(...lines.map(l => l.left)) : c.x
  const tr = lines.length ? Math.max(...lines.map(l => l.right)) : c.x + c.w
  const pct = p.lh ? p.lh / (fmax * lineFactor(fmaxFont)) : 1
  const brs = runs.filter(r => r.br).length
  const soft = !p.pre && n > brs + 1
  const sup = runs.filter(r => r.t !== undefined && r.va).reduce((s, r) => s + r.t.length * 0.4 * r.sz, 0)
  return { top, bottom, left: c.x, right: c.x + c.w, tl, tr, lh, font, fs, n, pct, soft, sup }
}

// A pre paragraph holding newlines becomes one paragraph per line.
function splitPreLines(p) {
  const runs = p._runs
  if (!runs.some(r => r.t !== undefined && r.t.includes('\n'))) return [p]
  const lines = [[]]
  for (const r of runs) {
    if (r.t === undefined) { lines[lines.length - 1].push(r); continue }
    r.t.split('\n').forEach((part, k) => { if (k) lines.push([]); if (part) lines[lines.length - 1].push({ ...r, t: part }) })
  }
  if (lines.length && !lines[lines.length - 1].length) lines.pop()
  const g = p._g
  return lines.map((lr, k) => ({ ...p, _runs: lr, _g: { ...g, top: g.top + k * g.lh, bottom: g.top + (k + 1) * g.lh, n: 1 }, id: p.id + '~' + k }))
}

const BULLET_AUTO = { decimal: 'arabicPeriod', 'lower-alpha': 'alphaLcPeriod', 'upper-alpha': 'alphaUcPeriod', 'lower-roman': 'romanLcPeriod', 'upper-roman': 'romanUcPeriod' }
function paraSpec(p, left, right, prevBottom) {
  const g = p._g
  const algn = { center: 'ctr', right: 'r', end: 'r', justify: 'just' }[p.align] || 'l'
  const marR = Math.max(0, right - g.right)
  const spec = { runs: [], algn, marL: r1(Math.max(0, g.left - left)), marR: algn !== 'l' || marR > 30 ? r1(marR) : 0, indent: 0 }
  const li = p.li
  if (li && li.type === 'pseudo') {
    spec.indent = r1(-(g.left - li.listLeft))
    spec.bullet = li.marker.auto ? { auto: li.marker.auto, start: li.index, color: hexc(li.marker.color) } : { char: li.marker.char, font: 'Arial', color: hexc(li.marker.color) }
  } else if (li && li.type && li.type !== 'none') {
    spec.indent = r1(-(g.left - li.listLeft))
    spec.bullet = BULLET_AUTO[li.type]
      ? { auto: BULLET_AUTO[li.type], start: li.index }
      : { char: { disc: '•', circle: 'o', square: '▪' }[li.type] || '•', font: li.type === 'circle' ? 'Courier New' : 'Arial', color: hexc(li.markerColor) }
  }
  spec.lnSpc = p.pre ? ['pts', r2(g.lh)] : ['pct', Math.round(g.pct * 100) * 1000]
  spec.spcBef = prevBottom != null ? r1(Math.max(0, g.top - prevBottom)) : 0
  for (const r of p._runs) {
    if (r.br) spec.runs.push({ br: true, sz: g.fs })
    else if (r.omml !== undefined) spec.runs.push(r)
    // PowerPoint shrinks super/subscript runs itself; keep the paragraph size.
    else spec.runs.push(r.va ? { ...r, sz: g.fs } : { ...r })
  }
  spec.endSz = g.fs
  spec.endFont = g.font
  return spec
}

// ---------------------------------------------------------------- scene ---

// SVG fills open paths too, so any filled path is an area, not a line
const svgClosed = it => ['rect', 'polygon', 'circle', 'ellipse'].includes(it.tag) || (it.tag === 'path' && !!it.fill && it.fill.a > 0.01 && it.fillOpacity > 0.01)
const degenerate = r => r.w < 0.6 && r.h < 0.6
const contains = (r, x, y, tol = 1) => r.x - tol <= x && x <= r.x + r.w + tol && r.y - tol <= y && y <= r.y + r.h + tol

// Facts across click states: boxes that may not hold their own text, and on which clicks each paragraph changes.
export function analyze(states) {
  const nonhost = new Set()
  for (const st of states) {
    for (const v of st.items) if (['box', 'svg', 'image', 'table'].includes(v.kind)) for (const b of v.boxChain || []) if (b !== v.id) nonhost.add(b)
    // svg groups with more than one real shape cannot host
    const groups = new Map()
    for (const it of st.items) if (it.kind === 'svg' && it.grp && !degenerate(it.bbox)) { if (!groups.has(it.grp)) groups.set(it.grp, []); groups.get(it.grp).push(it) }
    for (const its of groups.values()) if (its.length > 1) for (const i of its) nonhost.add('svg:' + i.id)
  }
  // Paragraphs that change on different clicks go into different text boxes, so that each box
  // changes as a whole and the unchanged ones stay single shapes.
  const seqs = new Map()
  states.forEach((st, i) => {
    for (const it of st.items) if (it.kind === 'para') {
      if (!seqs.has(it.id)) seqs.set(it.id, Array(states.length).fill(null))
      seqs.get(it.id)[i] = JSON.stringify(it.runs.map(r => [r.t, r.color && hexc(r.color), r.weight, r.s]))
    }
  })
  const changes = new Map()
  for (const [id, seq] of seqs) {
    const at = []
    if (seq[0] === null) at.push(-1)
    for (let i = 1; i < seq.length; i++) if (seq[i] !== seq[i - 1]) at.push(i)
    changes.set(id, at.join(','))
  }
  return { nonhost, changes }
}

export function buildState(st, facts) {
  const { nonhost, changes } = facts
  const items = st.items
  for (const it of items) if (it.kind === 'para') it._runs = normRuns(it)
  let paras = items.filter(it => it.kind === 'para' && (it.pre || it._runs.some(r => (r.t !== undefined && r.t.trim()) || r.omml !== undefined)))
  for (const p of paras) p._g = paraGeom(p)
  paras = paras.flatMap(p => p.pre ? splitPreLines(p) : [p])

  const svgs = items.filter(it => it.kind === 'svg' && !degenerate(it.bbox))
  const svgByGroup = new Map()
  for (const s of svgs) if (s.grp) { if (!svgByGroup.has(s.grp)) svgByGroup.set(s.grp, []); svgByGroup.get(s.grp).push(s) }

  const objs = new Map()
  const placeholders = new Map()
  paras = paras.filter(p => {
    if (!['title', 'ctrTitle', 'subTitle', 'trail'].includes(p.role)) return true
    if (!placeholders.has(p.role)) placeholders.set(p.role, [])
    placeholders.get(p.role).push(p)
    return false
  })

  // --- which shape each paragraph belongs to
  const hostOf = p => {
    const chain = p.boxChain || []
    if (p.svgGroup) {
      const grp = svgByGroup.get(p.svgGroup) || []
      const outer = grp[0]?.boxChain || []
      if (chain.length <= outer.length) {
        const g = p._g, cx = (g.tl + g.tr) / 2, cy = (g.top + g.bottom) / 2
        const cands = grp.filter(s => svgClosed(s) && contains(s.bbox, cx, cy, 2))
        if (cands.length) return 'svg:' + cands.reduce((a, b) => b.bbox.w * b.bbox.h > a.bbox.w * a.bbox.h ? b : a).id
        return null
      }
    }
    return chain.length ? chain[chain.length - 1] : null
  }
  const hosted = new Map()
  for (const p of paras) {
    p._host = hostOf(p)
    if (p._host && !nonhost.has(p._host)) { if (!hosted.has(p._host)) hosted.set(p._host, []); hosted.get(p._host).push(p) }
  }
  // a shape's paragraphs must stack vertically
  for (const [h, ps] of hosted) {
    ps.sort(byId)
    if (!ps.every((b, i) => i === 0 || b._g.top >= ps[i - 1]._g.bottom - 3)) { nonhost.add(h); hosted.delete(h) }
  }

  for (const it of items) {
    if (it.kind === 'box') for (const [k, o] of boxObjects(it, hosted.get(it.id))) objs.set(k, o)
    else if (it.kind === 'veil') {
      const r = it.rect
      objs.set('veil:' + it.id, { type: 'shape', x: r.x, y: r.y, w: r.w, h: r.h, geom: { prst: 'rect' }, fill: { color: hexc(it.back), alpha: Math.round((1 - it.op) * 1000) / 1000 }, line: null, text: null, order: [1e9, ...pathKey(it.id)], name: 'Dim Overlay' })
    } else if (it.kind === 'image') {
      objs.set('img:' + it.id, { type: 'pic', src: it.src, x: it.rect.x, y: it.rect.y, w: it.rect.w, h: it.rect.h, crop: it.crop, alpha: Math.round(it.op * 1000) / 1000, order: pathKey(it.id), name: 'Picture', descr: it.alt || '' })
    } else if (it.kind === 'snapshot') {
      objs.set('snap:' + it.id, { type: 'pic', src: 'snapshot:' + it.id, x: it.rect.x, y: it.rect.y, w: it.rect.w, h: it.rect.h, crop: null, alpha: 1, order: pathKey(it.id), name: 'Fallback Picture', descr: '' })
    } else if (it.kind === 'table') objs.set('tbl:' + it.id, tableObject(it))
    else if (it.kind === 'svgtext') { const o = svgTextObject(it); if (o) objs.set('stx:' + it.id, o) }
  }
  for (const s of svgs) for (const [k, o] of svgObjects(s, hosted.get('svg:' + s.id))) objs.set(k, o)

  // --- free text frames
  const unhosted = paras.filter(p => !(p._host && hosted.has(p._host))).sort(byId)
  const visuals = [...objs.values()].filter(o => ['shape', 'pic', 'table', 'cxn'].includes(o.type))
  const frames = []
  for (const p of unhosted) {
    const g = p._g, ctx = `${p._host}|${p.svgGroup}`
    const f = frames[frames.length - 1]
    if (f && f.ctx === ctx && canMerge(f, p, visuals, changes)) {
      f.paras.push(p); f.bottom = Math.max(f.bottom, g.bottom); f.left = Math.min(f.left, frameLeft(p)); f.right = Math.max(f.right, g.right)
    } else frames.push({ ctx, host: p._host, paras: [p], top: g.top, bottom: g.bottom, left: frameLeft(p), right: g.right, pre: p.pre })
  }
  for (const f of frames) {
    const o = frameObject(f)
    // CSS paints text above the block backgrounds of its stacking context
    const hk = f.host && !f.host.startsWith('svg:') ? pathKey(f.host) : []
    o.order = [...hk, 99999, ...pathKey(f.paras[0].id)]
    objs.set('txt:' + f.paras[0].id, o)
  }

  for (const [role, ps] of placeholders) {
    ps.sort(byId)
    const left = Math.min(...ps.map(frameLeft)), right = Math.max(...ps.map(q => q._g.right))
    let prev = null
    const specs = ps.map(q => { const s = paraSpec(q, left, right, prev); prev = q._g.bottom; return s })
    objs.set('ph:' + role, { type: 'ph', role, paras: specs, order: [role === 'trail' ? -2 : -1], x: left, y: ps[0]._g.top, w: right - left, h: ps[ps.length - 1]._g.bottom - ps[0]._g.top })
  }
  return objs
}

const frameLeft = p => p.li && p.li.type && p.li.type !== 'none' ? Math.min(p._g.left, p.li.listLeft) : p._g.left

function canMerge(f, p, visuals, changes) {
  const g = p._g, last = f.paras[f.paras.length - 1]
  if ((changes.get(last.id.split('~')[0]) || '') !== (changes.get(p.id.split('~')[0]) || '')) return false
  if (f.pre || p.pre) return !!(f.pre && p.pre && last._host === p._host && g.top >= f.bottom - 4 && g.top - f.bottom < g.lh * 1.5)
  if (g.top < f.bottom - 3) return false
  if (g.top - f.bottom > Math.max(30, 1.1 * g.lh)) return false
  const pl = frameLeft(p)
  if (pl < f.left - 2) return false
  if (g.right > f.right + 2 && Math.abs(pl - f.left) > 2) return false
  // alignment must agree for centred text, and centred lines must share one axis
  if ((p.align === 'center') !== (last.align === 'center')) return false
  if (p.align === 'center' && Math.abs((pl + g.right) / 2 - (f.left + f.right) / 2) > 3) return false
  // nothing visual may sit in the gap between them
  const gx0 = Math.max(f.left, pl), gx1 = Math.min(f.right, g.right), gy0 = f.bottom + 0.5, gy1 = g.top - 0.5
  if (gy1 > gy0) for (const o of visuals) {
    const [ox0, oy0, ox1, oy1] = o.type === 'cxn'
      ? [Math.min(o.x1, o.x2), Math.min(o.y1, o.y2), Math.max(o.x1, o.x2), Math.max(o.y1, o.y2)]
      : [o.x, o.y, o.x + o.w, o.y + o.h]
    if (ox1 <= gx0 || ox0 >= gx1 || oy1 <= gy0 || oy0 >= gy1) continue
    // an element that encloses both paragraphs is a background, not a separator
    if (ox0 <= f.left + 1 && ox1 >= g.right - 1 && oy0 <= f.top + 1 && oy1 >= g.bottom - 1) continue
    return false
  }
  return true
}

function frameObject(f) {
  const ps = f.paras
  let left = f.left
  const right = f.right
  let prev = null
  const specs = ps.map(p => { const s = paraSpec(p, left, right, prev); prev = p._g.bottom; return s })
  const allc = specs.every(s => s.algn === 'ctr'), allr = specs.every(s => s.algn === 'r')
  let w = right - left
  const sup = Math.max(...ps.map(p => p._g.sup))
  let slack = SLACK + 0.01 * w + sup
  const tl = Math.min(...ps.map(p => p._g.tl)), tr = Math.max(...ps.map(p => p._g.tr))
  let wrap = 'square'
  if (!f.pre && !ps.some(p => p._g.soft) && tr - tl < 420 && w - (tr - Math.min(tl, left)) < Math.max(24, 0.1 * w)) {
    // a short label that fills its column stays on one line, like a text box made by clicking
    wrap = 'none'
    slack = 4 + sup
    if (allc) { left = tl; w = tr - tl; for (const s of specs) s.marL = s.marR = 0 }
    else if (allr) { left = tl; w = right - tl }
    else { w = tr - left; for (const s of specs) s.marR = 0 }
  }
  if (f.pre) {
    // unwrapped code: the box hugs the longest line so no renderer re-centres it
    w = Math.max(...ps.map(p => p._g.tr)) - left
    slack = 4
    wrap = 'none'
  }
  const x = allc ? left - slack / 2 : allr ? left - slack : left
  return { type: 'shape', txBox: true, x, y: f.top, w: w + slack, h: f.bottom - f.top, geom: { prst: 'rect' }, fill: null, line: null,
    text: { paras: specs, ins: [0, 0, 0, 0], anchor: 't', wrap, autofit: 'sp' }, order: pathKey(ps[0].id), name: 'TextBox' }
}

function hostedText(rect, ps, { pre = false, centerHint = false, label = false, slanted = false, vcenter = false } = {}) {
  ps = [...ps].sort(byId)
  const tl = Math.min(...ps.map(frameLeft)), tr = Math.max(...ps.map(p => p._g.right))
  const ttl = Math.min(...ps.map(p => p._g.tl)), ttr = Math.max(...ps.map(p => p._g.tr))
  const top = ps[0]._g.top, bot = ps[ps.length - 1]._g.bottom
  const x0 = rect.x, x1 = rect.x + rect.w, y0 = rect.y, y1 = rect.y + rect.h
  let centred = ps.every(p => p.align === 'center')
  const gl = tl - x0, gr = x1 - tr, tgl = ttl - x0, tgr = x1 - ttr
  if (!centred && !pre && ps.every(p => p._g.n === 1 && !p.li) && (centerHint || (ps.length === 1 && tgl > 2 && Math.abs(tgl - tgr) < 1.5))) {
    centred = true
    for (const p of ps) p.align = 'center'
  }
  let lI, rI, left, right
  if (centred) {
    // symmetric insets that still leave room for the widest line
    const side = Math.max(0, Math.min(gl, gr, (rect.w - (ttr - ttl)) / 2 - SLACK))
    lI = rI = side; left = x0 + side; right = x1 - side
  } else {
    lI = Math.max(0, gl); rI = Math.max(0, gr - SLACK - 0.01 * rect.w); left = x0 + lI; right = x1 - rI
  }
  const gt = top - y0, gb = y1 - bot
  let anchor, tI, bI
  if (vcenter && Math.abs(gt - gb) <= Math.max(3, 0.25 * ps[0]._g.lh)) { anchor = 'ctr'; tI = bI = Math.max(0, Math.min(gt, gb, 4)) }
  else if (gt > gb + 3 && gb < 6) { anchor = 'b'; tI = 0; bI = Math.max(0, gb) }
  else { anchor = 't'; tI = Math.max(0, gt); bI = 0 }
  let prev = null
  const specs = ps.map(p => {
    const s = paraSpec(p, left, right, prev)
    prev = p._g.bottom
    if (s.algn === 'ctr' && !s.bullet) s.marL = s.marR = 0
    return s
  })
  let wrap = pre ? 'none' : 'square'
  if (!pre && !ps.some(p => p._g.soft) && (label || slanted || rect.w - lI - rI - (ttr - ttl) < 16)) wrap = 'none'
  if (slanted) lI = rI = 0
  return { paras: specs, ins: [lI, tI, rI, bI].map(r2), anchor, wrap, autofit: 'none' }
}

const lineSpec = (color, w, dash = null, alpha = 1, head = null, tail = null) => ({ color, w: r2(w), dash, alpha: Math.round(alpha * 1000) / 1000, head, tail })
const dashOf = style => ({ dashed: 'dash', dotted: 'sysDot' })[style] || null

function boxObjects(b, ps) {
  const out = new Map()
  const r = b.rect, bs = b.borders
  const vis = bs.filter(Boolean)
  const uniform = vis.length === 4 && vis.every(s => Math.abs(s.w - vis[0].w) < 0.3 && hexc(s.color) === hexc(vis[0].color) && s.style === vis[0].style)
  let fill = b.bg ? { color: hexc(b.bg), alpha: 1 } : null
  if (b.gradient) fill = { gradient: { angle: b.gradient.angle, stops: b.gradient.stops.map(s => ({ pos: s.pos, color: hexc(s.color) })) } }
  else if (b.bgImage) fill = patternFill(b.bgImage, b.op ?? 1) || fill
  const key = 'box:' + b.id, order = pathKey(b.id)
  let rect = r, line = null
  if (b.ring && (uniform || !vis.length)) {
    // the ring widens the outline outwards
    const w0 = uniform ? vis[0].w : 0, w = w0 + b.ring.w, out = b.ring.w
    rect = { x: r.x - out + w / 2, y: r.y - out + w / 2, w: r.w + 2 * out - w, h: r.h + 2 * out - w }
    line = lineSpec(hexc(uniform ? vis[0].color : b.ring.color), w, uniform ? dashOf(vis[0].style) : null)
  } else if (uniform) {
    const w = vis[0].w
    rect = { x: r.x + w / 2, y: r.y + w / 2, w: r.w - w, h: r.h - w }
    line = lineSpec(hexc(vis[0].color), w, dashOf(vis[0].style))
  }
  let geom = { prst: 'rect' }
  if (b.radius > 0.5) geom = { prst: 'roundRect', adj: Math.min(50000, Math.round(b.radius / Math.max(1, Math.min(rect.w, rect.h)) * 100000)) }
  const flexy = /flex|grid/.test(b.display || '')
  const shadow = b.shadow ? { color: hexc(b.shadow.color), alpha: b.shadow.alpha, dx: b.shadow.dx, dy: b.shadow.dy, blur: b.shadow.blur } : null
  if (fill || line || ps || shadow) out.set(key, {
    type: 'shape', x: rect.x, y: rect.y, w: rect.w, h: rect.h, geom, fill, line, shadow,
    text: ps ? hostedText(r, ps, { pre: b.tag === 'pre', centerHint: b.justify === 'center' && /flex/.test(b.display || ''), vcenter: b.alignItems === 'center' && flexy }) : null,
    order, name: geom.prst === 'rect' ? 'Rectangle' : 'Rounded Rectangle',
  })
  if (!uniform) bs.forEach((s, i) => {
    if (!s) return
    const w = s.w
    const [x1, y1, x2, y2] = [
      [r.x, r.y + w / 2, r.x + r.w, r.y + w / 2],
      [r.x + r.w - w / 2, r.y, r.x + r.w - w / 2, r.y + r.h],
      [r.x, r.y + r.h - w / 2, r.x + r.w, r.y + r.h - w / 2],
      [r.x + w / 2, r.y, r.x + w / 2, r.y + r.h],
    ][i]
    out.set(`${key}:side${i}`, { type: 'cxn', x1, y1, x2, y2, line: lineSpec(hexc(s.color), w, dashOf(s.style)), order: [...order, -1, i], name: 'Straight Connector' })
  })
  return out
}

function svgTextObject(it) {
  const p = { runs: it.runs, pre: false, whiteSpace: 'nowrap', fs: it.runs.find(r => r.size)?.size || 16 }
  const runs = normRuns(p)
  if (!runs.some(r => r.t && r.t.trim())) return null
  const algn = { middle: 'ctr', end: 'r' }[it.anchor] || 'l'
  const fs = runs.find(r => r.sz)?.sz || snap(p.fs)
  const r = it.rect, slack = 4
  const x = algn === 'ctr' ? r.x - slack / 2 : algn === 'r' ? r.x - slack : r.x
  return { type: 'shape', txBox: true, x, y: r.y, w: r.w + slack, h: r.h, geom: { prst: 'rect' }, fill: null, line: null,
    text: { paras: [{ runs, algn, marL: 0, marR: 0, indent: 0, lnSpc: ['pct', 100000], spcBef: 0, endSz: fs, endFont: runs[0]?.font }], ins: [0, 0, 0, 0], anchor: 't', wrap: 'none', autofit: 'sp' },
    order: pathKey(it.id), name: 'TextBox' }
}

function svgObjects(s, ps) {
  const out = new Map()
  const key = 'svg:' + s.id, M = s.M
  const sc = Math.sqrt(Math.abs(M[0] * M[3] - M[1] * M[2])) || 1
  const order = pathKey(s.id), op = s.op
  let fill = null
  if (s.fill && s.fill.a > 0.01 && s.fillOpacity > 0.01) fill = { color: hexc(s.fill), alpha: Math.round(s.fill.a * s.fillOpacity * op * 1000) / 1000 }
  let line = null
  const sw = s.nonScaling ? s.strokeWidth : s.strokeWidth * sc
  if (s.stroke && s.stroke.a > 0.01 && s.strokeOpacity > 0.01 && s.strokeWidth > 0) {
    let dash = null
    if (s.dash && !['none', '0', '0px'].includes(s.dash)) {
      const vals = (s.dash.match(/[\d.]+/g) || []).map(Number)
      if (vals.length && Math.max(...vals) > 0) dash = vals[0] <= 2.5 ? 'sysDot' : 'dash'
    }
    line = lineSpec(hexc(s.stroke), sw, dash, s.stroke.a * s.strokeOpacity * op)
  }
  const tag = s.tag, g = s.geom
  if (tag === 'image') {
    const [x0, y0] = apply(M, g.x, g.y), [x1, y1] = apply(M, g.x + g.w, g.y + g.h)
    out.set('img:' + s.id, { type: 'pic', src: g.href, x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0), crop: null, alpha: Math.round(op * 1000) / 1000, order, name: 'Picture', descr: '' })
    return out
  }
  let rect, geom, name
  if (tag === 'rect') {
    const [x0, y0] = apply(M, g.x, g.y), [x1, y1] = apply(M, g.x + g.w, g.y + g.h)
    rect = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) }
    const rx = (g.rx || 0) * sc
    geom = rx > 0.5 ? { prst: 'roundRect', adj: Math.min(50000, Math.round(rx / Math.max(1, Math.min(rect.w, rect.h)) * 100000)) } : { prst: 'rect' }
    name = 'Rectangle'
  } else if (tag === 'circle' || tag === 'ellipse') {
    const [x0, y0] = apply(M, g.cx - g.rx, g.cy - g.ry), [x1, y1] = apply(M, g.cx + g.rx, g.cy + g.ry)
    rect = { x: Math.min(x0, x1), y: Math.min(y0, y1), w: Math.abs(x1 - x0), h: Math.abs(y1 - y0) }
    geom = { prst: 'ellipse' }; name = 'Oval'
  } else {
    let subs, closed, pts
    if (['polygon', 'polyline', 'line'].includes(tag)) {
      pts = g.pts.map(p => apply(M, ...p))
      closed = tag === 'polygon'
      subs = [{ segs: [['M', pts[0]], ...pts.slice(1).map(p => ['L', p])], closed }]
    } else {
      subs = transformSubs(parsePath(g.d), M)
      closed = subs.some(sp => sp.closed) || !!fill
    }
    if (!subs.length || !subs[0].segs.length) return out
    let [bx0, by0, bx1, by1] = bbox(pointsOf(subs))
    rect = { x: bx0, y: by0, w: bx1 - bx0, h: by1 - by0 }
    if (!closed) {
      // open path: an edge. Straight → connector, otherwise a freeform with arrowheads.
      if (!line) return out
      const head = arrow(s.markerStart), tail = arrow(s.markerEnd)
      line.head = head; line.tail = tail
      const onlyLines = subs.length === 1 && subs[0].segs.every(seg => seg[0] === 'M' || seg[0] === 'L')
      let lp = onlyLines ? subs[0].segs.map(seg => seg[1]) : null
      if (lp) {
        lp = simplifyPolyline(lp)
        const ext = ((s.markerEnd || s.markerStart || {}).mw || 8) * sc * 0.5
        if (tail) lp[lp.length - 1] = extend(lp[lp.length - 2], lp[lp.length - 1], ext)
        if (head) lp[0] = extend(lp[1], lp[0], ext)
        if (lp.length === 2) {
          out.set(key, { type: 'cxn', x1: lp[0][0], y1: lp[0][1], x2: lp[1][0], y2: lp[1][1], line, order, name: head || tail ? 'Straight Arrow Connector' : 'Straight Connector' })
          return out
        }
        subs = [{ segs: [['M', lp[0]], ...lp.slice(1).map(p => ['L', p])], closed: false }]
      }
      ;[bx0, by0, bx1, by1] = bbox(pointsOf(subs))
      rect = { x: bx0, y: by0, w: Math.max(bx1 - bx0, 0.01), h: Math.max(by1 - by0, 0.01) }
      out.set(key, { type: 'shape', ...rect, geom: { cust: relSubs(subs, rect), open: true }, fill: null, line, text: null, order, name: 'Freeform' })
      return out
    }
    geom = null; name = 'Freeform'
    if (tag === 'polygon' && g.pts.length === 4) { geom = parallelogram(pts, rect); if (geom) name = 'Parallelogram' }
    if (!geom && tag === 'path' && s.grp && /[aA]/.test(g.d) && subs.length <= 2) { geom = { prst: 'can', adj: 25000 }; name = 'Cylinder' }
    if (!geom) geom = { cust: relSubs(subs, rect) }
  }
  const text = ps ? hostedText(rect, ps, { centerHint: true, label: true, slanted: ['parallelogram', 'can', 'hexagon', 'diamond'].includes(geom.prst), vcenter: true }) : null
  if (!fill && !line && !text) return out
  out.set(key, { type: 'shape', ...rect, geom, fill, line, text, order, name })
  return out
}

const arrow = m => m && m.type && m.type !== 'none' ? m.type : null
function extend(a, b, d) {
  const dx = b[0] - a[0], dy = b[1] - a[1], len = Math.hypot(dx, dy)
  return len < 1e-6 ? b : [b[0] + dx / len * d, b[1] + dy / len * d]
}
const relSubs = (subs, rect) => subs.map(sp => ({ segs: sp.segs.map(([k, ...ps]) => [k, ...ps.map(p => [Math.round((p[0] - rect.x) * 1000) / 1000, Math.round((p[1] - rect.y) * 1000) / 1000])]), closed: sp.closed }))

function parallelogram(pts, rect) {
  if (new Set(pts.map(p => r1(p[1]))).size !== 2) return null
  const top = pts.filter(p => Math.abs(p[1] - rect.y) < 0.6).sort((a, b) => a[0] - b[0])
  const bot = pts.filter(p => Math.abs(p[1] - (rect.y + rect.h)) < 0.6).sort((a, b) => a[0] - b[0])
  if (top.length !== 2 || bot.length !== 2) return null
  const offL = top[0][0] - bot[0][0], offR = top[1][0] - bot[1][0]
  if (Math.abs(offL - offR) > 1 || Math.abs(offL) < 1) return null
  return { prst: 'parallelogram', adj: Math.round(Math.abs(offL) / Math.min(rect.w, rect.h) * 100000), flipH: offL < 0 }
}

function tableObject(t) {
  const xs = [...new Set([...t.rows.flatMap(r => r.cells.map(c => r1(c.rect.x))), ...t.rows.flatMap(r => r.cells.filter(c => c.colSpan === 1).map(c => r1(c.rect.x + c.rect.w)))])].sort((a, b) => a - b)
  const cols = []
  for (const x of xs) if (!cols.length || x - cols[cols.length - 1] > 1.5) cols.push(x)
  // PowerPoint wraps text that exactly fills a cell; give such columns a few pixels more.
  const extra = cols.slice(1).map(() => 0)
  for (const r of t.rows) for (const c of r.cells) {
    if (c.colSpan !== 1) continue
    const i = cols.findIndex(x => Math.abs(x - c.rect.x) <= 1.5)
    const lines = c.paras.flatMap(p => p.lines || [])
    if (i < 0 || i >= extra.length || !lines.length) continue
    const room = c.rect.x + c.rect.w - c.pad[1] - Math.max(...lines.map(l => l.right))
    extra[i] = Math.max(extra[i], 4 - room)
  }
  let shift = 0
  for (let i = 0; i < extra.length; i++) { shift += Math.max(0, extra[i]); cols[i + 1] += shift }
  const rows = t.rows.map(r => ({
    h: r.rect.h,
    cells: r.cells.map(c => {
      for (const p of c.paras) { p._runs = normRuns(p); p._g = paraGeom(p) }
      const ps = c.paras.filter(p => p._runs.some(q => (q.t !== undefined && q.t.trim()) || q.omml !== undefined)).sort(byId)
      const left = c.rect.x + c.pad[3], right = c.rect.x + c.rect.w - c.pad[1]
      let prev = null
      const specs = ps.map(p => { const s = paraSpec(p, left, right, prev); prev = p._g.bottom; return s })
      return {
        paras: specs, fill: c.bg ? hexc(c.bg) : null, borders: c.borders.map(b => b ? lineSpec(hexc(b.color), b.w, dashOf(b.style)) : null),
        pad: c.pad.map(r2), anchor: { middle: 'ctr', bottom: 'b' }[c.valign] || 't', colSpan: c.colSpan,
      }
    }),
  }))
  return { type: 'table', x: cols[0], y: t.rect.y, w: cols[cols.length - 1] - cols[0], h: t.rect.h, cols: cols.slice(1).map((b, i) => b - cols[i]), rows, order: pathKey(t.id), name: 'Table' }
}

// Repeating stripe / grid gradients become PowerPoint pattern fills.
function patternFill(css, op) {
  const cols = []
  for (const m of css.matchAll(/rgba?\(([^)]+)\)|#([0-9a-fA-F]{6})/g)) {
    let c
    if (m[1]) {
      const v = m[1].trim().split(/[\s,/]+/).filter(Boolean).map(parseFloat)
      if ((v.length > 3 ? v[3] : 1) < 0.05) continue
      c = { r: v[0], g: v[1], b: v[2] }
    } else c = { r: parseInt(m[2].slice(0, 2), 16), g: parseInt(m[2].slice(2, 4), 16), b: parseInt(m[2].slice(4, 6), 16) }
    if (op < 0.999) c = { r: c.r * op + 255 * (1 - op), g: c.g * op + 255 * (1 - op), b: c.b * op + 255 * (1 - op) }
    cols.push(hexc(c))
  }
  const uniq = [...new Set(cols)]
  if (uniq.length < 2) return null
  const lum = h => [0, 2, 4].reduce((s, i) => s + parseInt(h.slice(i, i + 2), 16), 0)
  const light = uniq.reduce((a, b) => lum(b) > lum(a) ? b : a)
  if ((css.match(/repeating-linear-gradient/g) || []).length >= 2) return { pattern: 'lgGrid', fg: uniq.reduce((a, b) => lum(b) < lum(a) ? b : a), bg: light }
  const prst = css.includes('45deg') && !css.includes('-45deg') ? 'wdUpDiag' : (css.includes('135deg') || css.includes('-45deg')) ? 'wdDnDiag' : 'ltHorz'
  return { pattern: prst, fg: uniq.find(c => c !== light), bg: light }
}
