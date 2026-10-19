// SVG path parsing and affine helpers, in slide pixels.

const TOKEN = /[MmLlHhVvCcSsQqTtAaZz]|[-+]?(?:\d*\.\d+|\d+\.?)(?:[eE][-+]?\d+)?/g

export const apply = ([a, b, c, d, e, f], x, y) => [a * x + c * y + e, b * x + d * y + f]

function arcToCubics(x1, y1, rx, ry, phi, fa, fs, x2, y2) {
  if (rx === 0 || ry === 0) return [[[x1, y1], [x2, y2], [x2, y2]]]
  phi = phi * Math.PI / 180
  const cp = Math.cos(phi), sp = Math.sin(phi)
  const dx = (x1 - x2) / 2, dy = (y1 - y2) / 2
  const x1p = cp * dx + sp * dy, y1p = -sp * dx + cp * dy
  rx = Math.abs(rx); ry = Math.abs(ry)
  const lam = x1p ** 2 / rx ** 2 + y1p ** 2 / ry ** 2
  if (lam > 1) { rx *= Math.sqrt(lam); ry *= Math.sqrt(lam) }
  const num = rx ** 2 * ry ** 2 - rx ** 2 * y1p ** 2 - ry ** 2 * x1p ** 2
  const den = rx ** 2 * y1p ** 2 + ry ** 2 * x1p ** 2
  let co = den ? Math.sqrt(Math.max(0, num / den)) : 0
  if (fa === fs) co = -co
  const cxp = co * rx * y1p / ry, cyp = -co * ry * x1p / rx
  const cx = cp * cxp - sp * cyp + (x1 + x2) / 2, cy = sp * cxp + cp * cyp + (y1 + y2) / 2
  const ang = (ux, uy, vx, vy) => Math.atan2(ux * vy - uy * vx, ux * vx + uy * vy)
  const t1 = ang(1, 0, (x1p - cxp) / rx, (y1p - cyp) / ry)
  let dt = ang((x1p - cxp) / rx, (y1p - cyp) / ry, (-x1p - cxp) / rx, (-y1p - cyp) / ry)
  if (!fs && dt > 0) dt -= 2 * Math.PI
  else if (fs && dt < 0) dt += 2 * Math.PI
  const n = Math.max(1, Math.ceil(Math.abs(dt) / (Math.PI / 2) - 1e-9))
  const step = dt / n, out = []
  const pt = t => [cx + rx * Math.cos(t) * cp - ry * Math.sin(t) * sp, cy + rx * Math.cos(t) * sp + ry * Math.sin(t) * cp]
  const dv = t => [-rx * Math.sin(t) * cp - ry * Math.cos(t) * sp, -rx * Math.sin(t) * sp + ry * Math.cos(t) * cp]
  for (let i = 0; i < n; i++) {
    const a1 = t1 + i * step, a2 = t1 + (i + 1) * step, k = 4 / 3 * Math.tan((a2 - a1) / 4)
    const p1 = pt(a1), p2 = pt(a2), d1 = dv(a1), d2 = dv(a2)
    out.push([[p1[0] + k * d1[0], p1[1] + k * d1[1]], [p2[0] - k * d2[0], p2[1] - k * d2[1]], p2])
  }
  return out
}

// Subpaths of ['M', p] / ['L', p] / ['C', p1, p2, p] segments in user space.
export function parsePath(d) {
  const toks = d.match(TOKEN) || []
  const subs = []
  let i = 0, cur = null, x = 0, y = 0, sx = 0, sy = 0, cmd = null, lastCtrl = null, lastQ = null
  const num = () => parseFloat(toks[i++])
  while (i < toks.length) {
    const t = toks[i]
    if (/[A-Za-z]/.test(t)) {
      cmd = t; i++
      if (cmd === 'Z' || cmd === 'z') {
        if (cur) cur.closed = true
        x = sx; y = sy; lastCtrl = lastQ = null
        continue
      }
    } else if (cmd === null) { i++; continue }
    const rel = cmd === cmd.toLowerCase(), C = cmd.toUpperCase()
    if (C === 'M') {
      let nx = num(), ny = num()
      if (rel) { nx += x; ny += y }
      cur = { segs: [['M', [nx, ny]]], closed: false }
      subs.push(cur)
      x = sx = nx; y = sy = ny
      cmd = rel ? 'l' : 'L'
      lastCtrl = lastQ = null
    } else if (C === 'L') {
      let nx = num(), ny = num()
      if (rel) { nx += x; ny += y }
      cur.segs.push(['L', [nx, ny]]); x = nx; y = ny; lastCtrl = lastQ = null
    } else if (C === 'H') {
      let nx = num(); if (rel) nx += x
      cur.segs.push(['L', [nx, y]]); x = nx; lastCtrl = lastQ = null
    } else if (C === 'V') {
      let ny = num(); if (rel) ny += y
      cur.segs.push(['L', [x, ny]]); y = ny; lastCtrl = lastQ = null
    } else if (C === 'C') {
      let v = [num(), num(), num(), num(), num(), num()]
      if (rel) v = v.map((n, k) => n + (k % 2 ? y : x))
      cur.segs.push(['C', [v[0], v[1]], [v[2], v[3]], [v[4], v[5]]])
      lastCtrl = [v[2], v[3]]; lastQ = null; x = v[4]; y = v[5]
    } else if (C === 'S') {
      let v = [num(), num(), num(), num()]
      if (rel) v = v.map((n, k) => n + (k % 2 ? y : x))
      const c1 = lastCtrl ? [2 * x - lastCtrl[0], 2 * y - lastCtrl[1]] : [x, y]
      cur.segs.push(['C', c1, [v[0], v[1]], [v[2], v[3]]])
      lastCtrl = [v[0], v[1]]; lastQ = null; x = v[2]; y = v[3]
    } else if (C === 'Q') {
      let v = [num(), num(), num(), num()]
      if (rel) v = v.map((n, k) => n + (k % 2 ? y : x))
      const q = [v[0], v[1]]
      cur.segs.push(['C', [x + 2 / 3 * (q[0] - x), y + 2 / 3 * (q[1] - y)], [v[2] + 2 / 3 * (q[0] - v[2]), v[3] + 2 / 3 * (q[1] - v[3])], [v[2], v[3]]])
      lastQ = q; lastCtrl = null; x = v[2]; y = v[3]
    } else if (C === 'T') {
      let v = [num(), num()]
      if (rel) v = [v[0] + x, v[1] + y]
      const q = lastQ ? [2 * x - lastQ[0], 2 * y - lastQ[1]] : [x, y]
      cur.segs.push(['C', [x + 2 / 3 * (q[0] - x), y + 2 / 3 * (q[1] - y)], [v[0] + 2 / 3 * (q[0] - v[0]), v[1] + 2 / 3 * (q[1] - v[1])], v])
      lastQ = q; x = v[0]; y = v[1]
    } else if (C === 'A') {
      let [rx, ry, phi, fa, fs, nx, ny] = [num(), num(), num(), num(), num(), num(), num()]
      if (rel) { nx += x; ny += y }
      for (const [c1, c2, p] of arcToCubics(x, y, rx, ry, phi, Math.trunc(fa), Math.trunc(fs), nx, ny)) cur.segs.push(['C', c1, c2, p])
      x = nx; y = ny; lastCtrl = lastQ = null
    } else i++
  }
  return subs
}

export const transformSubs = (subs, M) => subs.map(s => ({ segs: s.segs.map(([k, ...ps]) => [k, ...ps.map(p => apply(M, ...p))]), closed: s.closed }))

export function* pointsOf(subs) {
  for (const s of subs) for (const [, ...ps] of s.segs) yield* ps
}

export function bbox(points) {
  let x0 = Infinity, y0 = Infinity, x1 = -Infinity, y1 = -Infinity
  for (const [x, y] of points) { x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y) }
  return [x0, y0, x1, y1]
}

// Drop interior points that lie on the line through their neighbours.
export function simplifyPolyline(pts, tol = 0.6) {
  if (pts.length <= 2) return pts
  const out = [pts[0]]
  for (let i = 1; i < pts.length - 1; i++) {
    const a = out[out.length - 1], b = pts[i], c = pts[i + 1]
    const cross = (b[0] - a[0]) * (c[1] - a[1]) - (b[1] - a[1]) * (c[0] - a[0])
    const len = Math.hypot(c[0] - a[0], c[1] - a[1]) || 1
    if (Math.abs(cross) / len > tol) out.push(b)
  }
  out.push(pts[pts.length - 1])
  return out
}
