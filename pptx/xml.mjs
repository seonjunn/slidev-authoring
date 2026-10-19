// DrawingML / PresentationML serialisation of object specs. One CSS pixel is 1/96 inch.
export const EMU = 9525
export const NS = {
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  p: 'http://schemas.openxmlformats.org/presentationml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
export const NSDECL = `xmlns:a="${NS.a}" xmlns:p="${NS.p}" xmlns:r="${NS.r}"`

export const E = px => Math.round(px * EMU)
// Hundredths of a point; values within rounding of a quarter point snap to it.
export function pts100(px) {
  const v = px * 75, q = v / 25
  return Math.abs(q - Math.round(q)) < 0.08 ? Math.round(q) * 25 : Math.round(v / 10) * 10
}
export const esc = t => String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')

export class Ctx {
  // linkRel(url) → relationship id; themeColors maps hex values to scheme colour names
  constructor({ linkRel = () => null, themeColors = {}, baseFont = 'Arial' } = {}) { this.linkRel = linkRel; this.themeColors = themeColors; this.baseFont = baseFont }
}

function clr(ctx, hex, alpha = 1) {
  const a = alpha != null && alpha < 0.999 ? `<a:alpha val="${Math.round(alpha * 100000)}"/>` : ''
  const scheme = ctx?.themeColors?.[hex]
  if (scheme && !a) return `<a:schemeClr val="${scheme}"/>`
  return a ? `<a:srgbClr val="${hex}">${a}</a:srgbClr>` : `<a:srgbClr val="${hex}"/>`
}
const solid = (ctx, hex, alpha = 1) => `<a:solidFill>${clr(ctx, hex, alpha)}</a:solidFill>`

function fillXml(ctx, f) {
  if (!f) return '<a:noFill/>'
  if (f.pattern) return `<a:pattFill prst="${f.pattern}"><a:fgClr><a:srgbClr val="${f.fg}"/></a:fgClr><a:bgClr><a:srgbClr val="${f.bg}"/></a:bgClr></a:pattFill>`
  if (f.gradient) {
    const ang = (((f.gradient.angle - 90) % 360) + 360) % 360
    const stops = f.gradient.stops.map(s => `<a:gs pos="${Math.round(s.pos * 100000)}"><a:srgbClr val="${s.color}"/></a:gs>`).join('')
    return `<a:gradFill rotWithShape="1"><a:gsLst>${stops}</a:gsLst><a:lin ang="${Math.round(ang * 60000)}" scaled="0"/></a:gradFill>`
  }
  return solid(ctx, f.color, f.alpha ?? 1)
}

function lnXml(ctx, line) {
  if (!line) return '<a:ln><a:noFill/></a:ln>'
  let inner = solid(ctx, line.color, line.alpha ?? 1)
  if (line.dash) inner += `<a:prstDash val="${line.dash}"/>`
  if (line.head) inner += `<a:headEnd type="${line.head}" w="med" len="med"/>`
  if (line.tail) inner += `<a:tailEnd type="${line.tail}" w="med" len="med"/>`
  return `<a:ln w="${Math.max(1, E(line.w))}">${inner}</a:ln>`
}

function effectXml(s) {
  if (!s) return ''
  const dist = Math.hypot(s.dx, s.dy), dir = ((Math.atan2(s.dy, s.dx) * 180 / Math.PI) + 360) % 360
  return `<a:effectLst><a:outerShdw blurRad="${E(s.blur)}" dist="${E(dist)}" dir="${Math.round(dir * 60000)}" algn="ctr" rotWithShape="0"><a:srgbClr val="${s.color}"><a:alpha val="${Math.round(s.alpha * 100000)}"/></a:srgbClr></a:outerShdw></a:effectLst>`
}

// ------------------------------------------------------------------ text --

function rprXml(r, ctx, tag = 'a:rPr', d = {}, mixed = false) {
  const at = ['lang="en-US"']
  if (r.sz !== undefined && (d.sz == null || Math.abs(r.sz - d.sz) > 0.05)) at.push(`sz="${pts100(r.sz)}"`)
  if (r.b && !d.b) at.push('b="1"')
  else if (!r.b && d.b && r.t !== undefined) at.push('b="0"')
  if (r.i) at.push('i="1"')
  if (r.u) at.push('u="sng"')
  else if (r.link) at.push('u="none"')
  if (r.s) at.push('strike="sngStrike"')
  if (r.va === 'super') at.push('baseline="30000"')
  else if (r.va === 'sub') at.push('baseline="-25000"')
  if (r.spc) at.push(`spc="${pts100(r.spc)}"`)
  if (r.caps) at.push('cap="all"')
  at.push('dirty="0"')
  let ch = ''
  if (r.color && r.color !== d.color) ch += solid(ctx, r.color)
  if (r.hl) ch += `<a:highlight><a:srgbClr val="${r.hl}"/></a:highlight>`
  if (r.font && r.font !== (d.font || ctx?.baseFont || 'Arial')) ch += `<a:latin typeface="${esc(r.font)}"/><a:cs typeface="${esc(r.font)}"/>`
  // beside runs in another font, name the theme font too, as PowerPoint does for "(Body)" text
  else if (mixed && r.t !== undefined) ch += '<a:latin typeface="+mn-lt"/><a:cs typeface="+mn-lt"/>'
  if (r.link && ctx?.linkRel) {
    const rid = ctx.linkRel(r.link)
    // Office 2019+ keeps the run's own colour for a link only with this extension
    if (rid) ch += `<a:hlinkClick r:id="${rid}"><a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}"><ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst></a:hlinkClick>`
  }
  return ch ? `<${tag} ${at.join(' ')}>${ch}</${tag}>` : `<${tag} ${at.join(' ')}/>`
}

export function paraXml(p, ctx, { ph = null, fallback = false, defaults = {}, mixed = false } = {}) {
  const at = []
  if (!ph) {
    if (p.marL) at.push(`marL="${E(p.marL)}"`)
    if (p.marR) at.push(`marR="${E(p.marR)}"`)
    if (p.indent) at.push(`indent="${E(p.indent)}"`)
  }
  if (p.algn !== (ph?.algn || 'l')) at.push(`algn="${p.algn}"`)
  let ch = ''
  if (!ph || ph.explicit) {
    const [kind, v] = p.lnSpc
    ch += kind === 'pct' ? `<a:lnSpc><a:spcPct val="${v}"/></a:lnSpc>` : `<a:lnSpc><a:spcPts val="${pts100(v)}"/></a:lnSpc>`
    if (p.spcBef) ch += `<a:spcBef><a:spcPts val="${pts100(p.spcBef)}"/></a:spcBef>`
  }
  if (!ph) {
    const b = p.bullet
    const buClr = b?.color ? `<a:buClr>${clr(ctx, b.color)}</a:buClr>` : ''
    if (b?.auto) ch += `${buClr}<a:buFont typeface="+mj-lt"/><a:buAutoNum type="${b.auto}"${b.start > 1 ? ` startAt="${b.start}"` : ''}/>`
    else if (b) ch += `${buClr}<a:buFont typeface="${b.font}"/><a:buChar char="${esc(b.char)}"/>`
  }
  const ppr = ch ? `<a:pPr ${at.join(' ')}>${ch}</a:pPr>` : at.length ? `<a:pPr ${at.join(' ')}/>` : ''
  let runs = ''
  for (const r of p.runs) {
    if (r.br) runs += `<a:br>${rprXml(r, ctx, 'a:rPr', defaults)}</a:br>`
    else if (r.omml !== undefined) {
      if (fallback || !r.omml) runs += `<a:r>${rprXml({ sz: r.sz, color: r.color, font: 'Cambria Math', i: true }, ctx, 'a:rPr', defaults)}<a:t>${esc(r.tex || '')}</a:t></a:r>`
      else runs += r.omml
    } else runs += `<a:r>${rprXml(r, ctx, 'a:rPr', defaults, mixed)}<a:t>${esc(r.t)}</a:t></a:r>`
  }
  const end = p.endSz ? { sz: p.endSz, font: p.endFont } : {}
  return `<a:p>${ppr}${runs}${rprXml(end, ctx, 'a:endParaRPr', defaults)}</a:p>`
}

export const hasMath = o => (o.text?.paras || o.paras || []).some(p => p.runs.some(r => r.omml !== undefined && r.omml))
export const mathOnly = o => { const runs = (o.text?.paras || []).flatMap(p => p.runs); return runs.length > 0 && runs.every(r => r.omml !== undefined) }

const mixedFonts = (paras, ctx) => paras.some(p => p.runs.some(r => r.font && r.font !== (ctx?.baseFont || 'Arial')))

function bodyXml(t, ctx, fallback = false) {
  const [l, tp, r, b] = t.ins
  const mixed = mixedFonts(t.paras, ctx)
  const fit = t.autofit === 'sp' ? '<a:spAutoFit/>' : '<a:noAutofit/>'
  return `<p:txBody><a:bodyPr wrap="${t.wrap}" lIns="${E(l)}" tIns="${E(tp)}" rIns="${E(r)}" bIns="${E(b)}" rtlCol="0" anchor="${t.anchor}">${fit}</a:bodyPr><a:lstStyle/>${t.paras.map(p => paraXml(p, ctx, { fallback, mixed })).join('')}</p:txBody>`
}

function geomXml(o) {
  const g = o.geom
  if (g.cust) {
    const w = Math.max(1, E(o.w)), h = Math.max(1, E(o.h))
    const pt = p => `<a:pt x="${E(p[0])}" y="${E(p[1])}"/>`
    const paths = g.cust.map(sp => {
      let body = ''
      for (const [k, ...ps] of sp.segs) body += k === 'M' ? `<a:moveTo>${pt(ps[0])}</a:moveTo>` : k === 'L' ? `<a:lnTo>${pt(ps[0])}</a:lnTo>` : `<a:cubicBezTo>${ps.map(pt).join('')}</a:cubicBezTo>`
      if (sp.closed) body += '<a:close/>'
      return `<a:path w="${w}" h="${h}"${g.open ? ' fill="none"' : ''}>${body}</a:path>`
    }).join('')
    return `<a:custGeom><a:avLst/><a:gdLst/><a:ahLst/><a:cxnLst/><a:rect l="l" t="t" r="r" b="b"/><a:pathLst>${paths}</a:pathLst></a:custGeom>`
  }
  return `<a:prstGeom prst="${g.prst}"><a:avLst>${g.adj !== undefined ? `<a:gd name="adj" fmla="val ${g.adj}"/>` : ''}</a:avLst></a:prstGeom>`
}

const xfrm = (o, flip = '') => `<a:xfrm${flip}><a:off x="${E(o.x)}" y="${E(o.y)}"/><a:ext cx="${Math.max(1, E(o.w))}" cy="${Math.max(1, E(o.h))}"/></a:xfrm>`

export function shapeXml(o, id, ctx, fallback = false) {
  const flip = o.geom.flipH ? ' flipH="1"' : ''
  const sppr = `<p:spPr>${xfrm(o, flip)}${geomXml(o)}${fillXml(ctx, o.fill)}${lnXml(ctx, o.line)}${effectXml(o.shadow)}</p:spPr>`
  const body = o.text ? bodyXml(o.text, ctx, fallback) : '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody>'
  const descr = o.descr ? ` descr="${esc(o.descr)}"` : ''
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(o.name + ' ' + id)}"${descr}/><p:cNvSpPr${o.txBox ? ' txBox="1"' : ''}/><p:nvPr/></p:nvSpPr>${sppr}${body}</p:sp>`
}

// A shape filled with a rendered equation, as PowerPoint writes for viewers without Office math.
export function pictureShapeXml(o, id, rid) {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${esc(o.name + ' ' + id)}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>`
    + `<p:spPr>${xfrm(o)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:blipFill><a:blip r:embed="${rid}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill><a:ln><a:noFill/></a:ln></p:spPr>`
    + '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>'
}

export function cxnXml(o, id, ctx) {
  let fl = ''
  if (o.x2 < o.x1) fl += ' flipH="1"'
  if (o.y2 < o.y1) fl += ' flipV="1"'
  const x = Math.min(o.x1, o.x2), y = Math.min(o.y1, o.y2), w = Math.abs(o.x2 - o.x1), h = Math.abs(o.y2 - o.y1)
  return `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${esc(o.name + ' ' + id)}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>`
    + `<p:spPr><a:xfrm${fl}><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom>${lnXml(ctx, o.line)}</p:spPr></p:cxnSp>`
}

export function picXml(o, id, rid) {
  const alpha = (o.alpha ?? 1) < 0.999 ? `<a:alphaModFix amt="${Math.round(o.alpha * 100000)}"/>` : ''
  const c = o.crop
  const crop = c ? `<a:srcRect l="${Math.round(c.l * 100000)}" t="${Math.round(c.t * 100000)}" r="${Math.round(c.r * 100000)}" b="${Math.round(c.b * 100000)}"/>` : ''
  const descr = o.descr ? ` descr="${esc(o.descr)}"` : ''
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(o.name + ' ' + id)}"${descr}/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="${rid}">${alpha}</a:blip>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`
    + `<p:spPr>${xfrm(o)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
}

export function tableXml(o, id, ctx) {
  const grid = o.cols.map(w => `<a:gridCol w="${E(w)}"/>`).join('')
  const rows = o.rows.map(r => {
    let tcs = ''
    for (const c of r.cells) {
      const paras = c.paras.map(p => paraXml(p, ctx, { mixed: mixedFonts(c.paras, ctx) })).join('') || '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
      const [pt, pr, pb, pl] = c.pad
      const lns = [['lnL', 3], ['lnR', 1], ['lnT', 0], ['lnB', 2]].map(([tag, i]) => {
        const b = c.borders[i]
        return b ? `<a:${tag} w="${Math.max(1, E(b.w))}" cap="flat" cmpd="sng" algn="ctr">${solid(ctx, b.color)}<a:prstDash val="${b.dash || 'solid'}"/></a:${tag}>` : `<a:${tag} w="0"><a:noFill/></a:${tag}>`
      }).join('')
      const fill = c.fill ? solid(ctx, c.fill) : '<a:noFill/>'
      const span = c.colSpan || 1
      tcs += `<a:tc${span > 1 ? ` gridSpan="${span}"` : ''}><a:txBody><a:bodyPr/><a:lstStyle/>${paras}</a:txBody><a:tcPr marL="${E(pl)}" marR="${E(pr)}" marT="${E(pt)}" marB="${E(pb)}" anchor="${c.anchor}">${lns}${fill}</a:tcPr></a:tc>`
      for (let k = 1; k < span; k++) tcs += '<a:tc hMerge="1"><a:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></a:txBody><a:tcPr/></a:tc>'
    }
    return `<a:tr h="${E(r.h)}">${tcs}</a:tr>`
  }).join('')
  return `<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="${id}" name="${esc(o.name + ' ' + id)}"/><p:cNvGraphicFramePr><a:graphicFrameLocks noGrp="1"/></p:cNvGraphicFramePr><p:nvPr/></p:nvGraphicFramePr>`
    + `<p:xfrm><a:off x="${E(o.x)}" y="${E(o.y)}"/><a:ext cx="${E(o.w)}" cy="${E(o.h)}"/></p:xfrm>`
    + `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl><a:tblPr firstRow="1"/><a:tblGrid>${grid}</a:tblGrid>${rows}</a:tbl></a:graphicData></a:graphic></p:graphicFrame>`
}

export const PH = {
  title: ['<p:ph type="title"/>', 'Title'],
  ctrTitle: ['<p:ph type="ctrTitle"/>', 'Title'],
  subTitle: ['<p:ph type="subTitle" idx="1"/>', 'Subtitle'],
  trail: ['<p:ph type="body" sz="quarter" idx="13"/>', 'Header Placeholder'],
}

// defaults: the layout's run and paragraph defaults; explicit: write position and spacing on the slide
export function phXml(o, id, ctx, defaults, { fallback = false, explicit = false } = {}) {
  const [phx, nm] = PH[o.role]
  const ph = { ...defaults, explicit }
  const paras = o.paras.map(p => paraXml(p, ctx, { ph, fallback, defaults })).join('')
  const sppr = explicit ? `<p:spPr>${xfrm(o)}</p:spPr>` : '<p:spPr/>'
  const body = explicit ? '<a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"><a:normAutofit/></a:bodyPr>' : '<a:bodyPr/>'
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${nm} ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phx}</p:nvPr></p:nvSpPr>${sppr}<p:txBody>${body}<a:lstStyle/>${paras}</p:txBody></p:sp>`
}

export const footerXml = (id, text) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Footer Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="ftr" sz="quarter" idx="11"/></p:nvPr></p:nvSpPr>`
  + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US" dirty="0"/><a:t>${esc(text)}</a:t></a:r><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`

export const slideNumberXml = (id, no) => `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Slide Number Placeholder ${id}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="12"/></p:nvPr></p:nvSpPr>`
  + `<p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US" dirty="0"/><a:t>${no}</a:t></a:fld><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>`

export const wrapMath = (choice, fallback) => '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">'
  + `<mc:Choice xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main" Requires="a14">${choice}</mc:Choice><mc:Fallback>${fallback}</mc:Fallback></mc:AlternateContent>`

// ------------------------------------------------------------- timing ----

export class Timing {
  constructor() { this.n = 2; this.grp = new Map() }
  nid() { return ++this.n }
  effect(e, nodeType) {
    const tgt = e.para !== undefined
      ? `<p:tgtEl><p:spTgt spid="${e.spid}"><p:txEl><p:pRg st="${e.para}" end="${e.para}"/></p:txEl></p:spTgt></p:tgtEl>`
      : `<p:tgtEl><p:spTgt spid="${e.spid}"/></p:tgtEl>`
    const g = (this.grp.get(e.spid) ?? -1) + 1
    this.grp.set(e.spid, g)
    e.grpId = g
    const head = (preset, cls, sub) => `<p:par><p:cTn id="${this.nid()}" presetID="${preset}" presetClass="${cls}" presetSubtype="${sub}" fill="hold" grpId="${g}" nodeType="${nodeType}"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>`
    const tail = '</p:childTnLst></p:cTn></p:par>'
    const set = (attr, val, dur = 1, override = '') => `<p:set><p:cBhvr${override}><p:cTn id="${this.nid()}" dur="${dur}" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${val}"/></p:to></p:set>`
    const anim = (attr, color, override = '') => `<p:animClr clrSpc="rgb" dir="cw"><p:cBhvr${override}><p:cTn id="${this.nid()}" dur="200" fill="hold"/>${tgt}<p:attrNameLst><p:attrName>${attr}</p:attrName></p:attrNameLst></p:cBhvr><p:to><a:srgbClr val="${color}"/></p:to></p:animClr>`
    switch (e.kind) {
      case 'appear': case 'disappear': {
        const h = head(1, e.kind === 'appear' ? 'entr' : 'exit', 0)
        return h + `<p:set><p:cBhvr><p:cTn id="${this.nid()}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>${tgt}<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr><p:to><p:strVal val="${e.kind === 'appear' ? 'visible' : 'hidden'}"/></p:to></p:set>` + tail
      }
      case 'fill': { const h = head(1, 'emph', 2); return h + anim('fillcolor', e.color, ' override="childStyle"') + set('fill.type', 'solid', 200) + set('fill.on', 'true', 200) + tail }
      case 'line': { const h = head(7, 'emph', 2); return h + anim('stroke.color', e.color) + set('stroke.on', 'true', 200) + tail }
      case 'font': { const h = head(3, 'emph', 2); return h + anim('style.color', e.color, ' override="childStyle"') + tail }
      // Bold Reveal: the weight stays changed until the slide ends
      case 'bold': { const h = head(15, 'emph', 0); return h + set('style.fontWeight', e.bold ? 'bold' : 'normal', 1, ' override="childStyle"') + tail }
    }
    throw new Error('unknown effect ' + e.kind)
  }
  // clicks[i]: effects started by click i + 1; kinds: shape id → 'sp' | 'pic' | 'cxn' | 'gf'
  xml(clicks, kinds) {
    let groups = ''
    for (const effects of clicks) {
      if (!effects.length) continue
      const a = this.nid(), b = this.nid()
      const inner = effects.map((e, i) => this.effect(e, i ? 'withEffect' : 'clickEffect')).join('')
      groups += `<p:par><p:cTn id="${a}" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst><p:par><p:cTn id="${b}" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>${inner}</p:childTnLst></p:cTn></p:par></p:childTnLst></p:cTn></p:par>`
    }
    if (!groups) return ''
    let bld = ''
    for (const e of clicks.flat()) {
      const k = kinds.get(e.spid)
      if (k === 'sp') bld += e.para !== undefined ? `<p:bldP spid="${e.spid}" grpId="${e.grpId}" build="p"/>` : `<p:bldP spid="${e.spid}" grpId="${e.grpId}" animBg="1"/>`
      else if (k === 'gf') bld += `<p:bldGraphic spid="${e.spid}" grpId="${e.grpId}"><p:bldAsOne/></p:bldGraphic>`
    }
    return '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>'
      + `<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>${groups}</p:childTnLst></p:cTn>`
      + '<p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>'
      + '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>'
      + `</p:childTnLst></p:cTn></p:par></p:tnLst>${bld ? `<p:bldLst>${bld}</p:bldLst>` : ''}</p:timing>`
  }
}
