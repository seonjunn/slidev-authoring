// These functions run in the browser through page.evaluate, so each one is self-contained.

// Record the PowerPoint font each element maps to, then lay the slide out with a font that has
// that font's metrics. Line breaks measured afterwards match what PowerPoint will draw.
export function prepareFonts({ no }) {
  const page = document.querySelector(`.slidev-page-${no}`)
  const OFFICE = {
    arial: 'Arial', helvetica: 'Arial', 'helvetica neue': 'Arial', georgia: 'Georgia', 'times new roman': 'Times New Roman',
    times: 'Times New Roman', 'courier new': 'Courier New', courier: 'Courier New', consolas: 'Consolas', calibri: 'Calibri',
    cambria: 'Cambria', verdana: 'Verdana', tahoma: 'Tahoma', 'trebuchet ms': 'Trebuchet MS', 'segoe ui': 'Segoe UI',
  }
  const MEASURE = {
    Arial: '"Liberation Sans", Arimo, Arial, sans-serif',
    Georgia: 'Georgia, Gelasio, serif',
    'Times New Roman': '"Liberation Serif", Tinos, "Times New Roman", serif',
    'Courier New': '"Liberation Mono", Cousine, "Courier New", monospace',
    Calibri: 'Carlito, Calibri, sans-serif',
    Cambria: 'Caladea, Cambria, serif',
  }
  const families = list => list.split(',').map(s => s.trim().replace(/^["']|["']$/g, '')).filter(Boolean)
  const map = list => {
    const fams = families(list), first = (fams[0] || '').toLowerCase(), all = fams.map(f => f.toLowerCase())
    if (first.startsWith('katex')) return null
    if (OFFICE[first]) return OFFICE[first]
    if (all.includes('monospace') || /mono|code|courier|menlo|consol/.test(first)) return 'Consolas'
    if ((all.includes('serif') && !all.includes('sans-serif')) || (/serif|georgia|times|garamond|gelasio/.test(first) && !/sans/.test(first))) return 'Georgia'
    return 'Arial'
  }
  const all = root => [...root.querySelectorAll('*')].flatMap(el => el.shadowRoot ? [el, ...all(el.shadowRoot)] : [el])
  const els = [page, ...all(page)].filter(el => !el.closest('.katex'))
  const plan = els.map(el => { const fam = getComputedStyle(el).fontFamily; return { el, fam, to: map(fam) } })
  const used = {}
  for (const { el, fam, to } of plan) {
    if (!to) continue
    el.setAttribute('data-pptx-font', to)
    const from = families(fam)[0] || ''
    const hasText = [...el.childNodes].some(n => n.nodeType === 3 && n.data.trim())
    if (hasText && from.toLowerCase() !== to.toLowerCase()) used[`${from} → ${to}`] = (used[`${from} → ${to}`] || 0) + 1
  }
  for (const { el, to } of plan) if (MEASURE[to]) el.style.setProperty('font-family', MEASURE[to], 'important')
  return used
}

// Read one click state of one slide into plain items: paragraphs with runs, boxes, SVG shapes,
// images, tables. Anything that cannot be represented becomes an issue.
export function extractSlide({ no, canvasWidth, skip, roles }) {
  const page = document.querySelector(`.slidev-page-${no}`)
  const pr = page.getBoundingClientRect()
  const S = pr.width / canvasWidth
  const rectOf = r => ({ x: (r.left - pr.left) / S, y: (r.top - pr.top) / S, w: r.width / S, h: r.height / S })
  const items = [], issues = []
  const ids = new WeakMap()
  const parentOf = el => el.parentElement || (el.parentNode instanceof ShadowRoot ? el.parentNode.host : null)
  function pathOf(el) {
    if (ids.has(el)) return ids.get(el)
    const p = parentOf(el)
    const s = !p || el === page ? 'P' : pathOf(p) + '/' + el.tagName.toLowerCase() + Array.prototype.indexOf.call(el.parentNode.children, el)
    ids.set(el, s)
    return s
  }
  function label(el) {
    const tag = el.tagName.toLowerCase()
    const cls = (el.getAttribute('class') || '').split(/\s+/).filter(Boolean)[0]
    const text = (el.textContent || '').replace(/\s+/g, ' ').trim().slice(0, 40)
    return `<${tag}${cls ? '.' + cls : ''}>` + (text ? ` "${text}"` : '')
  }
  function issue(el, kind, detail, action, severity = 'fix') {
    const r = el.getBoundingClientRect()
    issues.push({ kind, id: pathOf(el), rect: rectOf(r), label: label(el), detail, action, severity })
  }
  function parseColor(s) {
    if (!s || s === 'none' || s === 'transparent') return null
    let m = s.match(/rgba?\(([^)]+)\)/)
    if (m) {
      const p = m[1].split(/[\s,/]+/).filter(Boolean).map(v => v.endsWith('%') ? parseFloat(v) * 2.55 : parseFloat(v))
      return { r: p[0], g: p[1], b: p[2], a: p.length > 3 ? (m[1].includes('%') && p[3] > 1 ? p[3] / 255 : p[3]) : 1 }
    }
    m = s.match(/color\(srgb ([^)]+)\)/)
    if (m) { const p = m[1].split(/[\s/]+/).filter(Boolean).map(parseFloat); return { r: p[0] * 255, g: p[1] * 255, b: p[2] * 255, a: p.length > 3 ? p[3] : 1 } }
    return null
  }
  const layoutEl = page.querySelector('.slidev-layout') || page
  const pageBg = [layoutEl, page, page.parentElement, document.body].map(e => e && parseColor(getComputedStyle(e).backgroundColor)).find(c => c && c.a > 0.5) || { r: 255, g: 255, b: 255, a: 1 }
  function backdropOf(el) {
    for (let e = el.parentElement; e && e !== page; e = e.parentElement) {
      if (e instanceof SVGElement) continue
      const c = parseColor(getComputedStyle(e).backgroundColor)
      if (c && c.a > 0.5) return c
    }
    return pageBg
  }
  function blend(c, alpha, back) {
    if (!c) return null
    const a = c.a * alpha
    if (a >= 0.999) return { r: c.r, g: c.g, b: c.b, a: 1 }
    return { r: c.r * a + back.r * (1 - a), g: c.g * a + back.g * (1 - a), b: c.b * a + back.b * (1 - a), a: 1 }
  }
  function fontScaleOf(el) {
    const svg = el.closest('svg')
    if (svg) { const m = svg.getScreenCTM(); return m ? Math.hypot(m.a, m.b) / S : 1 }
    if (el.offsetWidth > 20) { const f = el.getBoundingClientRect().width / el.offsetWidth / S; if (Math.abs(f - 1) > 0.03) return f }
    return 1
  }
  const fontOf = el => { for (let e = el; e; e = parentOf(e)) if (e.hasAttribute('data-pptx-font')) return e.getAttribute('data-pptx-font'); return 'Arial' }

  // ---------- text runs ----------
  function inlineBg(pe, block) {
    for (let e = pe; e && e !== block; e = e.parentElement) {
      const cs = getComputedStyle(e)
      if (cs.display.startsWith('inline')) {
        const c = parseColor(cs.backgroundColor)
        if (c && c.a > 0.05) return c
      }
    }
    return null
  }
  function vAlignOf(pe, block) {
    for (let e = pe; e && e !== block; e = e.parentElement) {
      const t = e.tagName.toLowerCase(), va = getComputedStyle(e).verticalAlign
      if (va === 'super' || t === 'sup') return 'super'
      if (va === 'sub' || t === 'sub') return 'sub'
      if (/^-?[\d.]+px$/.test(va)) { const v = parseFloat(va); if (v > 1) return 'super'; if (v < -1) return 'sub' }
    }
    return null
  }
  function opacityBetween(pe, block) {
    let o = 1
    for (let e = pe; e && e !== block; e = e.parentElement) o *= parseFloat(getComputedStyle(e).opacity)
    return o
  }
  function runStyle(pe, block, fscale, op, back) {
    const cs = getComputedStyle(pe)
    const deco = cs.textDecorationLine || ''
    return {
      font: fontOf(pe), size: parseFloat(cs.fontSize) * fscale, weight: parseInt(cs.fontWeight), italic: cs.fontStyle === 'italic',
      color: blend(parseColor(cs.color), op * opacityBetween(pe, block), back), va: vAlignOf(pe, block),
      u: deco.includes('underline'), s: deco.includes('line-through'), hl: inlineBg(pe, block),
      ls: cs.letterSpacing === 'normal' ? 0 : parseFloat(cs.letterSpacing) * fscale, caps: cs.textTransform === 'uppercase',
      ws: cs.whiteSpace, link: pe.closest('a')?.getAttribute('href') || null,
    }
  }
  function collectRuns(node, block, runs, fscale, op, back) {
    if (node.nodeType === 3) {
      const pe = node.parentElement
      if (getComputedStyle(pe).visibility === 'hidden') return
      runs.push({ t: node.data, ...runStyle(pe, block, fscale, op, back) })
      return
    }
    if (node.nodeType !== 1) return
    const cs = getComputedStyle(node)
    if (cs.display === 'none' || parseFloat(cs.opacity) < 0.03) return
    const tag = node.tagName.toLowerCase()
    if (tag === 'br') { runs.push({ br: true }); return }
    if (node.classList.contains('katex')) {
      const math = node.querySelector('math')
      runs.push({
        math: math ? math.outerHTML : '', tex: node.querySelector('annotation')?.textContent || '', display: !!node.closest('.katex-display'),
        rect: rectOf(node.getBoundingClientRect()), mathId: pathOf(node), ...runStyle(node, block, fscale, op, back),
      })
      return
    }
    if (tag === 'svg' || tag === 'img') return
    for (const ch of node.childNodes) collectRuns(ch, block, runs, fscale, op, back)
  }
  function textRects(nodes) {
    const out = []
    const visit = n => {
      if (n.nodeType === 3) {
        if (!n.data.trim() && !(n.parentElement && getComputedStyle(n.parentElement).whiteSpace.startsWith('pre'))) return
        const r = document.createRange(); r.selectNodeContents(n)
        out.push(...r.getClientRects())
        return
      }
      if (n.nodeType !== 1) return
      const cs = getComputedStyle(n)
      if (cs.display === 'none' || cs.visibility === 'hidden') return
      if (n.classList.contains('katex')) { out.push(n.getBoundingClientRect()); return }
      for (const c of n.childNodes) visit(c)
    }
    nodes.forEach(visit)
    return out
  }
  function lineBoxes(nodes) {
    const rects = textRects(nodes).filter(r => r.width > 0.1 && r.height > 0.1).sort((a, b) => a.top - b.top)
    const lines = []
    for (const r of rects) {
      const L = lines.find(l => Math.min(l.bottom, r.bottom) - Math.max(l.top, r.top) > 0.5 * Math.min(l.bottom - l.top, r.height))
      if (L) { L.top = Math.min(L.top, r.top); L.bottom = Math.max(L.bottom, r.bottom); L.left = Math.min(L.left, r.left); L.right = Math.max(L.right, r.right); L.hs.push(r.height) }
      else lines.push({ top: r.top, bottom: r.bottom, left: r.left, right: r.right, hs: [r.height] })
    }
    return lines
  }
  // A list item whose marker is drawn with ::before (a dot, a short symbol, or a counter).
  function pseudoMarker(li) {
    if (!li || li.tagName.toLowerCase() !== 'li') return null
    const ps = getComputedStyle(li, '::before'), content = ps.content
    if (!content || content === 'none' || content === 'normal' || ps.display === 'none') return null
    const left = ps.position === 'absolute' ? li.getBoundingClientRect().left + parseFloat(ps.left || '0') * S : null
    if (/counter\(/.test(content)) return { auto: 'arabicPlain', color: parseColor(ps.color), left }
    const text = content.match(/^"(.*)"$/)?.[1]
    if (text === undefined) return null
    if (!text.trim()) { const bg = parseColor(ps.backgroundColor); return bg && bg.a > 0.05 ? { char: '▪', color: bg, left } : null }
    return [...text.trim()].length <= 2 ? { char: text.trim(), color: parseColor(ps.color), left } : null
  }
  function flushParagraph(block, nodes, ctx) {
    if (!nodes.length) return
    const fscale = fontScaleOf(block)
    const back = backdropOf(nodes[0].nodeType === 1 ? nodes[0] : nodes[0].parentElement)
    const runs = []
    for (const n of nodes) collectRuns(n, block, runs, fscale, ctx.op ?? 1, back)
    const textAll = runs.map(r => r.t || (r.math ? 'M' : '')).join('')
    const pre = !!ctx.inPre
    if (!pre && !textAll.trim()) return
    if (!runs.some(r => r.t !== undefined || r.math)) return
    const lines = lineBoxes(nodes)
    const cs = getComputedStyle(block)
    const br = block.getBoundingClientRect()
    const bl = parseFloat(cs.borderLeftWidth) + parseFloat(cs.paddingLeft), brr = parseFloat(cs.borderRightWidth) + parseFloat(cs.paddingRight)
    const bt = parseFloat(cs.borderTopWidth) + parseFloat(cs.paddingTop), bb = parseFloat(cs.borderBottomWidth) + parseFloat(cs.paddingBottom)
    const sc = fscale * S
    const content = { left: br.left + bl * sc, right: br.right - brr * sc, top: br.top + bt * sc, bottom: br.bottom - bb * sc }
    const lh = cs.lineHeight === 'normal' ? null : parseFloat(cs.lineHeight) * fscale
    const isItem = el => el && el.tagName.toLowerCase() === 'li' && (getComputedStyle(el).display === 'list-item' || pseudoMarker(el))
    // only the item's first paragraph carries the marker
    const leading = (parent, first) => [...parent.childNodes].slice(0, [...parent.childNodes].indexOf(first)).every(n => n.nodeType === 8 || (n.nodeType === 3 && !n.data.trim()))
    let liEl = null
    if (isItem(block)) { if (leading(block, nodes[0])) liEl = block }
    else {
      // <li><p>text</p></li>: the first block inside the item carries the marker
      const par = block.parentElement
      if (isItem(par) && leading(par, block) && leading(block, nodes[0])) liEl = par
    }
    let li = null
    const marker = liEl && pseudoMarker(liEl)
    if (marker) {
      const list = liEl.parentElement, idx = [...list.children].filter(c => c.tagName.toLowerCase() === 'li').indexOf(liEl)
      li = { type: 'pseudo', marker: { auto: marker.auto, char: marker.char, color: marker.color }, index: idx + 1,
        listLeft: ((marker.left ?? liEl.getBoundingClientRect().left) - pr.left) / S }
    } else if (liEl) {
      const lcs = getComputedStyle(liEl), list = liEl.parentElement
      const listCs = getComputedStyle(list), lr = list.getBoundingClientRect()
      const idx = [...list.children].filter(c => c.tagName.toLowerCase() === 'li').indexOf(liEl)
      li = {
        type: lcs.listStyleType, index: idx + parseInt(list.getAttribute('start') || '1'),
        listLeft: (lr.left + parseFloat(listCs.borderLeftWidth) * sc - pr.left) / S,
        markerColor: parseColor(getComputedStyle(liEl, '::marker').color) || parseColor(lcs.color),
      }
    }
    const onlyChild = [...block.childNodes].every(n => nodes.includes(n) || (n.nodeType === 3 && !n.data.trim()) || n.nodeType === 8)
    const writing = cs.writingMode
    if (writing && writing.startsWith('vertical')) issue(block, 'vertical-text', '세로쓰기 글은 가로로 들어갑니다.', '가로로 배치')
    if (cs.textShadow && cs.textShadow !== 'none') issue(block, 'text-shadow', '글자 그림자는 옮기지 않았습니다.', '생략', 'approx')
    const item = {
      kind: 'para', id: onlyChild ? pathOf(block) : pathOf(block) + '#' + [...block.childNodes].indexOf(nodes[0]),
      runs, align: cs.textAlign, fs: parseFloat(cs.fontSize) * fscale, lh, pre, li,
      lines: lines.map(l => ({ top: (l.top - pr.top) / S, bottom: (l.bottom - pr.top) / S, left: (l.left - pr.left) / S, right: (l.right - pr.left) / S, maxh: Math.max(...l.hs) / S })),
      content: { x: (content.left - pr.left) / S, y: (content.top - pr.top) / S, w: (content.right - content.left) / S, h: (content.bottom - content.top) / S },
      onlyChild, whiteSpace: cs.whiteSpace, boxChain: ctx.boxChain, svgGroup: ctx.svgGroup || null, role: ctx.role || null,
    }
    if (ctx.capture) ctx.capture.push(item); else items.push(item)
  }

  // ---------- CSS effects that PowerPoint lacks or draws differently ----------
  function checkStyle(el, cs) {
    if (cs.transform && cs.transform !== 'none') {
      const m = cs.transform.match(/matrix\(([^)]+)\)/)
      const v = m ? m[1].split(',').map(Number) : null
      if (!v || Math.abs(v[1]) > 1e-3 || Math.abs(v[2]) > 1e-3) issue(el, 'transform', '회전·기울임은 옮기지 못해 회전 전 모양으로 들어갑니다.', '회전 없이 배치')
    }
    for (const pseudo of ['::before', '::after']) {
      if (pseudo === '::before' && pseudoMarker(el)) { issue(el, 'list-marker', 'CSS로 그린 목록 기호를 PowerPoint 글머리표·번호로 바꿨습니다.', '글머리표로 근사', 'approx'); continue }
      const ps = getComputedStyle(el, pseudo)
      const content = ps.content
      if (!content || content === 'none' || content === 'normal' || ps.display === 'none') continue
      const visible = content !== '""' || (parseColor(ps.backgroundColor)?.a > 0.01) || parseFloat(ps.borderTopWidth) > 0 || parseFloat(ps.borderLeftWidth) > 0
      if (visible) issue(el, 'pseudo', `CSS ${pseudo}로 그린 내용(${content.slice(0, 30)})은 옮기지 못했습니다.`, '생략')
    }
    const bgImage = cs.backgroundImage
    if (bgImage && bgImage !== 'none' && /url\(/.test(bgImage)) issue(el, 'background-image', '배경 그림은 옮기지 못했습니다.', '생략')
    if (bgImage && /radial-gradient|conic-gradient/.test(bgImage)) issue(el, 'gradient', '원형·원뿔 그라데이션은 단색으로 바꿨습니다.', '단색으로 근사', 'approx')
    if (cs.filter && cs.filter !== 'none') issue(el, 'filter', `CSS filter(${cs.filter.slice(0, 30)})는 적용하지 않았습니다.`, '생략', 'approx')
    if (cs.clipPath && cs.clipPath !== 'none') issue(el, 'clip-path', '잘라내기(clip-path)는 적용하지 않았습니다.', '자르지 않고 배치', 'approx')
    if (cs.mixBlendMode && cs.mixBlendMode !== 'normal') issue(el, 'blend', `혼합 모드(${cs.mixBlendMode})는 반투명으로 근사했습니다.`, '반투명으로 근사', 'approx')
  }
  function shadowOf(el, cs) {
    const v = cs.boxShadow
    if (!v || v === 'none') return null
    const parts = v.split(/,(?![^(]*\))/)
    if (parts.some(p => /inset/.test(p))) { issue(el, 'shadow', '안쪽 그림자(inset)는 옮기지 못했습니다.', '생략'); return null }
    if (parts.length > 1) issue(el, 'shadow', '그림자가 여러 겹이라 첫 번째만 옮겼습니다.', '첫 그림자만', 'approx')
    const color = parseColor(parts[0].match(/rgba?\([^)]*\)/)?.[0])
    const nums = parts[0].replace(/rgba?\([^)]*\)/, '').match(/-?[\d.]+px/g)?.map(parseFloat) || []
    if (!color || color.a < 0.01) return null
    const [dx = 0, dy = 0, blur = 0, spread = 0] = nums
    // a spread-only shadow is a ring drawn outside the border
    if (!dx && !dy && !blur && spread > 0 && parts.length === 1) return { ring: spread, color }
    if (spread) issue(el, 'shadow', '그림자 퍼짐(spread)은 옮기지 못했습니다.', '퍼짐 없이 근사', 'approx')
    return { color, dx, dy, blur }
  }
  function gradientOf(css) {
    const m = css.match(/^linear-gradient\((.*)\)$/s)
    if (!m) return null
    const args = m[1].split(/,(?![^(]*\))/).map(s => s.trim())
    let angle = 180
    if (/deg$/.test(args[0])) angle = parseFloat(args.shift())
    else if (/^to /.test(args[0])) {
      const dir = args.shift()
      angle = { 'to top': 0, 'to right': 90, 'to bottom': 180, 'to left': 270, 'to top right': 45, 'to right top': 45, 'to bottom right': 135, 'to right bottom': 135, 'to bottom left': 225, 'to left bottom': 225, 'to top left': 315, 'to left top': 315 }[dir] ?? 180
    }
    const stops = args.map(a => { const c = parseColor(a.match(/(rgba?|color)\([^)]*\)/)?.[0]); const p = a.match(/([\d.]+)%/); return c && { color: c, pos: p ? parseFloat(p[1]) / 100 : null } }).filter(Boolean)
    if (stops.length < 2) return null
    stops.forEach((s, i) => { if (s.pos == null) s.pos = i / (stops.length - 1) })
    return { angle, stops }
  }

  // ---------- boxes ----------
  const SIDES = ['Top', 'Right', 'Bottom', 'Left']
  function boxVisual(el, cs, op, ctx) {
    if (el.classList.contains('slidev-layout') || el === page) return null
    const back = backdropOf(el)
    let bgc = parseColor(cs.backgroundColor)
    const fscale = fontScaleOf(el)
    const sides = SIDES.map(s => ({ w: parseFloat(cs[`border${s}Width`]) * fscale, style: cs[`border${s}Style`], color: parseColor(cs[`border${s}Color`]) }))
    let bgImage = cs.backgroundImage && cs.backgroundImage !== 'none' && cs.backgroundImage.includes('gradient') ? cs.backgroundImage : null
    const gradient = bgImage && !bgImage.includes('repeating') ? gradientOf(bgImage) : null
    if (bgImage && !gradient && !bgImage.includes('repeating')) {
      // radial and conic gradients: the first colour stands in for the whole
      const first = parseColor(bgImage.match(/(rgba?|color)\([^)]*\)/)?.[0])
      bgImage = null
      if (first && !(bgc && bgc.a > 0.01)) bgc = first
    }
    const shadow = shadowOf(el, cs)
    const hasBg = (bgc && bgc.a > 0.01) || !!bgImage
    const vis = sides.map(s => s.w > 0.2 && s.style !== 'none' && s.style !== 'hidden' && s.color && s.color.a > 0.01)
    if (!hasBg && !vis.some(Boolean) && !shadow) return null
    const r = el.getBoundingClientRect()
    if (r.width < 0.5 || r.height < 0.5) return null
    const id = pathOf(el)
    items.push({
      kind: 'box', id, rect: rectOf(r), tag: el.tagName.toLowerCase(),
      bg: bgc && bgc.a > 0.01 ? blend(bgc, op, back) : null, bgImage: gradient ? null : bgImage, op,
      gradient: gradient && { angle: gradient.angle, stops: gradient.stops.map(s => ({ pos: s.pos, color: blend(s.color, op, back) })) },
      shadow: shadow && !shadow.ring ? { ...shadow, dx: shadow.dx * fscale, dy: shadow.dy * fscale, blur: shadow.blur * fscale, alpha: shadow.color.a * op } : null,
      ring: shadow?.ring ? { w: shadow.ring * fscale, color: blend(shadow.color, op, back) } : null,
      borders: sides.map((s, i) => vis[i] ? { w: s.w, style: s.style, color: blend(s.color, op, back) } : null),
      radius: parseFloat(cs.borderTopLeftRadius) * fscale || 0,
      boxChain: ctx.boxChain, display: cs.display, justify: cs.justifyContent, alignItems: cs.alignItems,
    })
    return id
  }

  // ---------- images and things that can only be pictures ----------
  function addImg(el, cs, ctx, op) {
    const r = el.getBoundingClientRect()
    if (r.width < 1 || r.height < 1) return
    const nw = el.naturalWidth, nh = el.naturalHeight
    if (!nw || !nh) { issue(el, 'image-failed', `그림을 불러오지 못했습니다(${(el.getAttribute('src') || '').slice(0, 60)}).`, '생략'); return }
    let rect = rectOf(r), crop = null
    const fit = cs.objectFit
    if (fit === 'contain' || fit === 'scale-down') {
      const s = Math.min(rect.w / nw, rect.h / nh), w = nw * s, h = nh * s
      rect = { x: rect.x + (rect.w - w) / 2, y: rect.y + (rect.h - h) / 2, w, h }
    } else if (fit === 'cover') {
      const s = Math.max(rect.w / nw, rect.h / nh), w = nw * s, h = nh * s
      crop = { l: (w - rect.w) / 2 / w, r: (w - rect.w) / 2 / w, t: (h - rect.h) / 2 / h, b: (h - rect.h) / 2 / h }
    }
    items.push({ kind: 'image', id: pathOf(el), src: el.currentSrc || el.src, rect, crop, op, alt: el.alt || '' })
  }
  function addSnapshot(el, why) {
    const r = el.getBoundingClientRect()
    if (r.width < 1 || r.height < 1) return
    items.push({ kind: 'snapshot', id: pathOf(el), rect: rectOf(r), page: { x: r.left, y: r.top, w: r.width, h: r.height } })
    issue(el, 'snapshot', `${why}은(는) 편집할 수 없는 그림으로 넣었습니다.`, '그림으로 대체')
  }

  // ---------- tables ----------
  function addTable(tbl, ctx, op) {
    const rows = [], back = backdropOf(tbl)
    for (const tr of tbl.rows) {
      const trcs = getComputedStyle(tr)
      if (trcs.display === 'none') continue
      const trBg = parseColor(trcs.backgroundColor), cells = []
      for (const td of tr.cells) {
        const tcs = getComputedStyle(td), capture = []
        processChildren(td, { ...ctx, op, boxChain: [], capture })
        const fscale = fontScaleOf(td), bg = parseColor(tcs.backgroundColor)
        cells.push({
          id: pathOf(td), rect: rectOf(td.getBoundingClientRect()), colSpan: td.colSpan, rowSpan: td.rowSpan,
          bg: blend(bg && bg.a > 0.01 ? bg : (trBg && trBg.a > 0.01 ? trBg : null), op, back),
          borders: SIDES.map(s => { const w = parseFloat(tcs[`border${s}Width`]), c = parseColor(tcs[`border${s}Color`]); return w > 0.2 && tcs[`border${s}Style`] !== 'none' && c && c.a > 0.01 ? { w: w * fscale, color: blend(c, op, back), style: tcs[`border${s}Style`] } : null }),
          pad: SIDES.map(s => parseFloat(tcs[`padding${s}`]) * fscale), valign: tcs.verticalAlign, paras: capture,
        })
        if (td.rowSpan > 1) issue(td, 'table-rowspan', '세로로 합친 표 칸은 합치지 않은 칸으로 들어갑니다.', '칸 나눔', 'approx')
      }
      rows.push({ rect: rectOf(tr.getBoundingClientRect()), cells })
    }
    items.push({ kind: 'table', id: pathOf(tbl), rect: rectOf(tbl.getBoundingClientRect()), rows })
  }

  // ---------- svg ----------
  function markerInfo(ref, svg) {
    if (!ref || ref === 'none') return null
    const m = ref.match(/url\("?#([^")]+)"?\)/)
    if (!m) return null
    const mk = svg.ownerDocument.getElementById(m[1]) || svg.querySelector(`[id="${m[1]}"]`)
    if (!mk) return { type: 'triangle', mw: 8 }
    const id = m[1].toLowerCase()
    return { type: id.includes('circle') ? 'oval' : id.includes('cross') ? 'none' : 'triangle', mw: parseFloat(mk.getAttribute('markerWidth') || '8') }
  }
  function walkSvg(svg, ctx) {
    const sop = ctx.op * parseFloat(getComputedStyle(svg).opacity)
    if (sop < 0.03) return
    const rec = (node, g) => {
      for (const ch of node.children) {
        const tag = ch.tagName.toLowerCase()
        if (['style', 'defs', 'marker', 'clippath', 'lineargradient', 'radialgradient', 'title', 'desc', 'symbol', 'mask', 'pattern', 'filter', 'metadata'].includes(tag)) continue
        const cs = getComputedStyle(ch)
        if (cs.display === 'none') continue
        const op = g.op * parseFloat(cs.opacity)
        if (op < 0.03) continue
        if (cs.clipPath && cs.clipPath !== 'none' || cs.mask && cs.mask !== 'none') issue(ch, 'svg-clip', 'SVG 잘라내기·마스크는 적용하지 않았습니다.', '자르지 않고 배치', 'approx')
        if (tag === 'g' || tag === 'a' || tag === 'switch') {
          const cl = ch.getAttribute('class') || ''
          const grp = g.grp || (/\b(node|edgeLabel|cluster)\b/.test(cl) ? pathOf(ch) : null)
          rec(ch, { op, grp }); continue
        }
        if (tag === 'foreignobject') {
          for (const c2 of ch.children) walk(c2, { ...ctx, op, svgGroup: g.grp, boxChain: ctx.boxChain })
          continue
        }
        if (tag === 'svg') { rec(ch, g); continue }
        if (cs.visibility === 'hidden') continue
        const m = ch.getScreenCTM()
        if (!m) continue
        const M = [m.a / S, m.b / S, m.c / S, m.d / S, (m.e - pr.left) / S, (m.f - pr.top) / S]
        const A = n => ch.getAttribute(n)
        for (const prop of ['fill', 'stroke']) if (/url\(/.test(cs[prop])) issue(ch, 'svg-paint', `SVG ${prop === 'fill' ? '채우기' : '선'}의 그라데이션·패턴은 옮기지 못했습니다.`, '생략')
        let geom = null
        if (tag === 'rect') geom = { x: ch.x.baseVal.value, y: ch.y.baseVal.value, w: ch.width.baseVal.value, h: ch.height.baseVal.value, rx: parseFloat(cs.rx) || parseFloat(A('rx') || '0') || 0 }
        else if (tag === 'circle') geom = { cx: ch.cx.baseVal.value, cy: ch.cy.baseVal.value, rx: ch.r.baseVal.value, ry: ch.r.baseVal.value }
        else if (tag === 'ellipse') geom = { cx: ch.cx.baseVal.value, cy: ch.cy.baseVal.value, rx: ch.rx.baseVal.value, ry: ch.ry.baseVal.value }
        else if (tag === 'line') geom = { pts: [[ch.x1.baseVal.value, ch.y1.baseVal.value], [ch.x2.baseVal.value, ch.y2.baseVal.value]] }
        else if (tag === 'polyline' || tag === 'polygon') geom = { pts: [...ch.points].map(p => [p.x, p.y]) }
        else if (tag === 'path') geom = { d: A('d') || '' }
        else if (tag === 'image') geom = { x: ch.x.baseVal.value, y: ch.y.baseVal.value, w: ch.width.baseVal.value, h: ch.height.baseVal.value, href: new URL(A('href') || A('xlink:href'), location.href).href }
        else if (tag === 'text') {
          if (ch.querySelector('textPath')) { issue(ch, 'svg-text-path', '곡선을 따라가는 SVG 글은 곧은 글로 들어갑니다.', '곧게 배치', 'approx') }
          const runs = []
          collectRuns(ch, ch, runs, 1, op, pageBg)
          for (const r of runs) if (r.size) r.size = parseFloat(getComputedStyle(ch).fontSize) * Math.hypot(M[0], M[1])
          items.push({ kind: 'svgtext', id: pathOf(ch), runs, rect: rectOf(ch.getBoundingClientRect()), anchor: cs.textAnchor, grp: g.grp })
          continue
        }
        else { issue(ch, 'svg-element', `SVG <${tag}> 요소는 옮기지 못했습니다.`, '생략'); continue }
        items.push({
          kind: 'svg', tag, id: pathOf(ch), M, geom, bbox: rectOf(ch.getBoundingClientRect()), grp: g.grp,
          fill: /url\(/.test(cs.fill) ? null : parseColor(cs.fill), fillOpacity: parseFloat(cs.fillOpacity),
          stroke: /url\(/.test(cs.stroke) ? null : parseColor(cs.stroke), strokeOpacity: parseFloat(cs.strokeOpacity),
          strokeWidth: parseFloat(cs.strokeWidth), dash: cs.strokeDasharray, op, nonScaling: cs.vectorEffect === 'non-scaling-stroke',
          markerStart: markerInfo(cs.markerStart, svg), markerEnd: markerInfo(cs.markerEnd, svg), boxChain: ctx.boxChain,
        })
        if (cs.mixBlendMode !== 'normal' && !ch.classList.contains('focus-border')) issue(ch, 'blend', `혼합 모드(${cs.mixBlendMode})는 반투명으로 근사했습니다.`, '반투명으로 근사', 'approx')
      }
    }
    rec(svg, { op: sop, grp: null })
  }

  // ---------- html walk ----------
  function isInline(el, cs) {
    const d = cs.display
    if (d === 'inline' || d === 'contents') return true
    if (d === 'inline-block' || d === 'inline-flex' || d === 'inline-grid') return !el.querySelector('div,p,section,ul,ol,table,pre,h1,h2,h3,img,svg,canvas,video')
    return false
  }
  function processChildren(el, ctx) {
    let buf = []
    const flush = () => { if (buf.length) flushParagraph(el, buf, ctx); buf = [] }
    for (const ch of el.childNodes) {
      if (ch.nodeType === 3) { if (!(el instanceof ShadowRoot)) buf.push(ch); continue }
      if (ch.nodeType !== 1) continue
      const ccs = getComputedStyle(ch)
      if (ccs.display === 'none') continue
      const t = ch.tagName.toLowerCase()
      if (t === 'br' || (!['svg', 'img', 'canvas', 'video', 'iframe', 'object', 'embed'].includes(t) && (ch.classList.contains('katex') || isInline(ch, ccs)))) { buf.push(ch); continue }
      flush()
      walk(ch, ctx)
    }
    flush()
  }
  const SKIP = ['script', 'style', 'template', 'noscript']
  const state = { titleSeen: false }
  function roleOf(el, ctx) {
    for (const rule of roles) {
      if (!el.matches(rule.match)) continue
      if (rule.within && !el.closest(rule.within)) continue
      if (rule.first && state.titleSeen) continue
      return rule.role
    }
    return ctx.role
  }
  function walk(el, ctx) {
    const cs = getComputedStyle(el)
    if (cs.display === 'none') return
    const tag = el.tagName.toLowerCase()
    if (SKIP.includes(tag) || skip.some(sel => el.matches(sel))) return
    const own = parseFloat(cs.opacity)
    if (ctx.op * own < 0.03) return
    // A dimmed element keeps its own colours under a translucent overlay in the backdrop colour,
    // so its text stays one editable shape across click states.
    const veiled = own < 0.999 && !ctx.capture && tag !== 'svg'
    if (veiled) items.push({ kind: 'veil', id: pathOf(el), rect: rectOf(el.getBoundingClientRect()), op: own, back: backdropOf(el) })
    const op = veiled ? ctx.op : ctx.op * own
    if (!ctx.capture) checkStyle(el, cs)
    if (tag === 'svg') return walkSvg(el, { ...ctx, op })
    if (tag === 'img') return addImg(el, cs, ctx, op)
    if (tag === 'table') return addTable(el, ctx, op)
    if (['canvas', 'video', 'iframe', 'object', 'embed'].includes(tag)) return addSnapshot(el, `<${tag}>`)
    if (['button', 'input', 'select', 'textarea'].includes(tag) && !ctx.capture) issue(el, 'interactive', '버튼·입력 칸은 동작하지 않는 도형이 됩니다.', '정적 도형으로 배치', 'approx')
    const role = roleOf(el, ctx)
    if (role === 'title') state.titleSeen = true
    const myBox = cs.visibility !== 'hidden' && !ctx.capture ? boxVisual(el, cs, op, ctx) : null
    const inPre = ctx.inPre || tag === 'pre'
    const childCtx = { ...ctx, op, role, inPre, boxChain: myBox && !ctx.inPre ? [...ctx.boxChain, myBox] : ctx.boxChain }
    // components such as Slidev's Mermaid render into a shadow root
    if (el.shadowRoot) {
      if (el.shadowRoot.querySelector('slot')) issue(el, 'shadow-slot', 'shadow DOM의 slot 내용은 옮기지 못했을 수 있습니다.', '확인 필요', 'approx')
      processChildren(el.shadowRoot, childCtx)
    }
    else processChildren(el, childCtx)
  }
  if (layoutEl !== page) {
    const lcs = getComputedStyle(layoutEl)
    if (lcs.backgroundImage && lcs.backgroundImage !== 'none') issue(layoutEl, 'background-image', '슬라이드 배경 그림·그라데이션은 옮기지 못했습니다.', '생략')
  }
  walk(layoutEl, { op: 1, boxChain: [], role: null })
  // Global layers (slide-top, slide-bottom) sit beside the layout element.
  for (const el of page.children) if (el !== layoutEl && !el.contains(layoutEl)) walk(el, { op: 1, boxChain: [], role: null })
  return { pageBg, items, issues }
}

// mathml2omml output → the OMML PowerPoint writes inside a14:m, with Cambria Math run properties.
export function fixOmml({ omml, sz, color, display }) {
  const M = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
  const A = 'http://schemas.openxmlformats.org/drawingml/2006/main'
  const W = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
  const esc = t => t.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  const unesc = t => t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&amp;/g, '&')
  // mathml2omml leaves '<' and '&' unescaped inside m:t
  omml = omml.replace(/(<m:t(?:\s[^>]*)?>)([\s\S]*?)(<\/m:t>)/g, (_, a, t, b) => a + esc(unesc(t)) + b)
  const doc = new DOMParser().parseFromString(omml, 'application/xml')
  if (doc.querySelector('parsererror')) throw new Error('OMML parse error')
  const root = doc.documentElement
  const all = (ns, name) => [...root.getElementsByTagNameNS(ns, name)]
  const kids = (el, name) => [...el.children].filter(c => c.namespaceURI === M && c.localName === name)
  for (const el of all(W, '*')) el.remove()
  for (const sty of all(M, 'sty')) if (!['p', 'b', 'i', 'bi'].includes(sty.getAttributeNS(M, 'val') || sty.getAttribute('m:val'))) sty.remove()
  for (const rpr of all(M, 'rPr')) if (!rpr.children.length) rpr.remove()
  for (const r of all(M, 'r')) { const t = kids(r, 't')[0]; if (!t || !t.textContent.trim()) r.remove() }
  const FUNCS = new Set(['min', 'max', 'log', 'exp', 'sin', 'cos', 'tan', 'lim', 'argmin', 'argmax'])
  const arpr = italic => {
    const e = doc.createElementNS(A, 'a:rPr')
    e.setAttribute('lang', 'en-US'); e.setAttribute('sz', String(sz)); e.setAttribute('i', italic ? '1' : '0')
    if (color) { const f = doc.createElementNS(A, 'a:solidFill'), c = doc.createElementNS(A, 'a:srgbClr'); c.setAttribute('val', color); f.append(c); e.append(f) }
    const lat = doc.createElementNS(A, 'a:latin')
    lat.setAttribute('typeface', 'Cambria Math'); lat.setAttribute('panose', '02040503050406030204'); lat.setAttribute('pitchFamily', '18'); lat.setAttribute('charset', '0')
    e.append(lat)
    return e
  }
  for (const r of all(M, 'r')) {
    const t = kids(r, 't')[0]
    if (t && FUNCS.has(t.textContent.replace(/[⁡ ]/g, '')) && !kids(r, 'rPr').length) {
      const rpr = doc.createElementNS(M, 'm:rPr'), sty = doc.createElementNS(M, 'm:sty')
      sty.setAttributeNS(M, 'm:val', 'p'); rpr.append(sty); r.prepend(rpr)
    }
    const mrpr = kids(r, 'rPr')[0]
    const sty = mrpr && kids(mrpr, 'sty')[0]
    let italic = !(sty && ['p', 'b'].includes(sty.getAttributeNS(M, 'val') || sty.getAttribute('m:val')))
    if (t && !/[A-Za-z]/.test(t.textContent)) italic = false
    if (mrpr) mrpr.after(arpr(italic)); else r.prepend(arpr(italic))
  }
  for (const cp of all(M, 'ctrlPr')) if (!cp.children.length) cp.append(arpr(true))
  // n-ary operators whose base was left empty take the following factors
  for (const nary of all(M, 'nary')) {
    const e = kids(nary, 'e')[0]
    if (!e || e.children.length) continue
    let sib = nary.nextElementSibling
    while (sib) {
      if (sib.localName === 'r') { const t = kids(sib, 't')[0]; if (t && ['=', '+', '−', '-', '<', '>', '≈', '≤', '≥', ','].includes(t.textContent.trim())) break }
      const next = sib.nextElementSibling
      e.append(sib)
      sib = next
    }
  }
  const inner = [...root.children].map(c => new XMLSerializer().serializeToString(c)).join('')
  return display
    ? `<a14:m><m:oMathPara xmlns:m="${M}"><m:oMathParaPr><m:jc m:val="centerGroup"/></m:oMathParaPr><m:oMath>${inner}</m:oMath></m:oMathPara></a14:m>`
    : `<a14:m><m:oMath xmlns:m="${M}">${inner}</m:oMath></a14:m>`
}

// Bytes of an image the PowerPoint package cannot hold (SVG, WebP …) → PNG.
export async function rasterize({ url, w, h }) {
  const img = new Image()
  img.src = url
  await img.decode()
  const scale = 2
  const cw = Math.max(1, Math.round((img.naturalWidth || w) * scale)), ch = Math.max(1, Math.round((img.naturalHeight || h) * scale))
  const c = document.createElement('canvas')
  c.width = cw; c.height = ch
  c.getContext('2d').drawImage(img, 0, 0, cw, ch)
  return c.toDataURL('image/png')
}

// Where do two renderings of one slide differ? Both are downscaled and blurred so that
// antialiasing and hinting do not count; masked rectangles (slide px) are ignored.
export async function compareImages({ a, b, masks, width, height }) {
  const load = async src => { const i = new Image(); i.src = src; await i.decode(); return i }
  const [ia, ib] = await Promise.all([load(a), load(b)])
  const W = 192, H = Math.round(192 * height / width)
  const pixels = img => {
    const c = document.createElement('canvas'); c.width = W; c.height = H
    const x = c.getContext('2d', { willReadFrequently: true })
    x.fillStyle = '#fff'; x.fillRect(0, 0, W, H)
    x.imageSmoothingQuality = 'high'
    x.filter = 'blur(2px)'
    x.drawImage(img, 0, 0, W, H)
    x.fillStyle = '#fff'
    for (const m of masks) x.fillRect(m.x * W / width - 2, m.y * H / height - 2, m.w * W / width + 4, m.h * H / height + 4)
    return x.getImageData(0, 0, W, H).data
  }
  const pa = pixels(ia), pb = pixels(ib)
  let n = 0, x0 = W, y0 = H, x1 = -1, y1 = -1
  // A pixel differs only if no neighbour in the other image matches it: shifts of a few
  // pixels between two text renderers do not count, missing or extra content does.
  const diff = (i, j) => 0.299 * Math.abs(pa[i] - pb[j]) + 0.587 * Math.abs(pa[i + 1] - pb[j + 1]) + 0.114 * Math.abs(pa[i + 2] - pb[j + 2])
  for (let y = 0; y < H; y++) for (let x = 0; x < W; x++) {
    const i = (y * W + x) * 4
    let d = Infinity
    for (let dy = -1; dy <= 1 && d > 14; dy++) for (let dx = -1; dx <= 1 && d > 14; dx++) {
      const yy = y + dy, xx = x + dx
      if (yy < 0 || yy >= H || xx < 0 || xx >= W) continue
      d = Math.min(d, diff(i, (yy * W + xx) * 4))
    }
    if (d > 14) { n++; x0 = Math.min(x0, x); y0 = Math.min(y0, y); x1 = Math.max(x1, x); y1 = Math.max(y1, y) }
  }
  const box = n ? { x: x0 * width / W, y: y0 * height / H, w: (x1 - x0 + 1) * width / W, h: (y1 - y0 + 1) * height / H } : null
  return { score: n / (W * H) * 100, box }
}

// Original | converted, cropped to one region, for the report.
export async function sideBySide({ a, b, box, width, height }) {
  const load = async src => { const i = new Image(); i.src = src; await i.decode(); return i }
  const [ia, ib] = await Promise.all([load(a), load(b)])
  const pad = 24
  const x = Math.max(0, box.x - pad), y = Math.max(0, box.y - pad)
  const w = Math.min(width - x, box.w + 2 * pad), h = Math.min(height - y, box.h + 2 * pad)
  const scale = Math.min(2, 900 / w)
  const cw = Math.round(w * scale), ch = Math.round(h * scale), gap = 12
  const c = document.createElement('canvas'); c.width = cw * 2 + gap; c.height = ch
  const g = c.getContext('2d')
  g.fillStyle = '#fff'; g.fillRect(0, 0, c.width, c.height)
  g.drawImage(ia, x * ia.width / width, y * ia.height / height, w * ia.width / width, h * ia.height / height, 0, 0, cw, ch)
  g.drawImage(ib, x * ib.width / width, y * ib.height / height, w * ib.width / width, h * ib.height / height, cw + gap, 0, cw, ch)
  return c.toDataURL('image/png')
}

// Well-formedness check for generated XML parts before they are packaged.
export function xmlErrors(parts) {
  const out = []
  for (const [name, xml] of parts) {
    const doc = new DOMParser().parseFromString(xml, 'application/xml')
    const err = doc.querySelector('parsererror')
    if (err) out.push(`${name}: ${err.textContent.slice(0, 200)}`)
  }
  return out
}
