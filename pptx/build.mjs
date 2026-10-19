// Extracted click states → PowerPoint package parts.
import { createHash, randomUUID } from 'node:crypto'
import { analyze, buildState, hexc } from './scene.mjs'
import { planSlide, simulate } from './plan.mjs'
import * as X from './xml.mjs'
import { modeOf, notesMasterXml, notesSlideXml } from './profiles.mjs'

const XMLH = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
const CT = 'application/vnd.openxmlformats-officedocument.presentationml'
const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
export const texKey = tex => createHash('sha1').update(tex).digest('hex').slice(0, 16)

const relsXml = rels => `${XMLH}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">`
  + rels.map(r => `<Relationship Id="${r.id}" Type="${r.type}" Target="${X.esc(r.target)}"${r.external ? ' TargetMode="External"' : ''}/>`).join('') + '</Relationships>'

class Media {
  constructor() { this.byHash = new Map(); this.files = [] }
  add({ data, ext }) {
    const h = createHash('sha1').update(data).digest('hex')
    if (!this.byHash.has(h)) {
      const name = `media/image${this.files.length + 1}.${ext}`
      this.byHash.set(h, name)
      this.files.push([`ppt/${name}`, data])
    }
    return '../' + this.byHash.get(h)
  }
}

// deck: { title, author, canvas, profile, footer, themeImages, slides, images, equations }
// mode 'deck' animates every slide; mode 'states' writes one slide per click state at its final look.
export function buildPackage(deck, { mode = 'deck' } = {}) {
  const { profile, canvas } = deck
  const media = new Media()

  // ---- objects per click state, and the slide kind they need
  const slides = deck.slides.map(s => {
    const facts = analyze(s.states)
    const perState = s.states.map(st => buildState(st, facts))
    let kind = profile.kindOf(s.layout)
    if (profile.name === 'generic' && !perState.some(o => o.has('ph:title'))) kind = 'blank'
    return { ...s, perState, kind }
  })

  // ---- layout defaults from what the deck shows
  const firstRun = o => { const r = o.paras.flatMap(p => p.runs).find(r => r.t !== undefined); return r && { sz: Math.round(r.sz * 1000) / 1000, b: !!r.b, color: r.color, font: r.font } }
  const observed = (kind, role) => modeOf(slides.filter(s => s.kind === kind).map(s => s.perState.at(-1).get('ph:' + role)).filter(Boolean).map(firstRun).filter(Boolean))
  const defaults = profile.defaults(observed)
  const titleRects = slides.filter(s => s.kind === 'content').map(s => s.perState.at(-1).get('ph:title')).filter(Boolean).map(o => [o.x, o.y, o.w, o.h].map(v => Math.round(v)))
  const titleBox = modeOf(titleRects) || [40, 30, canvas.w - 80, 60]
  const weight = new Map()
  for (const s of slides) for (const o of s.perState.at(-1).values()) for (const p of o.text?.paras || o.paras || []) for (const r of p.runs)
    if (r.font && r.t && !['Consolas', 'Courier New'].includes(r.font)) weight.set(r.font, (weight.get(r.font) || 0) + r.t.length)
  const font = [...weight.entries()].reduce((a, b) => b[1] > a[1] ? b : a, ['Arial', 0])[0]

  // ---- master and layouts
  const files = []
  const kinds = profile.kinds
  const layoutIds = kinds.map((k, i) => `<p:sldLayoutId id="${2147483649 + i}" r:id="rId${i + 1}"/>`).join('')
  files.push(['ppt/slideMasters/slideMaster1.xml', profile.master({ layoutIds, defaults, footer: deck.footer, size: canvas, titleBox })])
  files.push(['ppt/slideMasters/_rels/slideMaster1.xml.rels', relsXml([...kinds.map((k, i) => ({ id: `rId${i + 1}`, type: `${REL}/slideLayout`, target: `../slideLayouts/slideLayout${i + 1}.xml` })), { id: `rId${kinds.length + 1}`, type: `${REL}/theme`, target: '../theme/theme1.xml' }])])
  kinds.forEach((kind, i) => {
    const rels = [{ id: 'rId1', type: `${REL}/slideMaster`, target: '../slideMasters/slideMaster1.xml' }], rids = {}
    for (const name of profile.layoutImages[kind] || []) {
      const img = deck.themeImages?.[name]
      if (!img) continue
      rids[name] = `rId${rels.length + 1}`
      rels.push({ id: rids[name], type: `${REL}/image`, target: media.add(img) })
    }
    files.push([`ppt/slideLayouts/slideLayout${i + 1}.xml`, profile.layout(kind, { defaults, footer: deck.footer, rids, titleBox })])
    files.push([`ppt/slideLayouts/_rels/slideLayout${i + 1}.xml.rels`, relsXml(rels)])
  })
  files.push(['ppt/theme/theme1.xml', profile.theme({ font })])
  files.push(['ppt/theme/theme2.xml', profile.theme({ font })])
  files.push(['ppt/notesMasters/notesMaster1.xml', notesMasterXml()])
  files.push(['ppt/notesMasters/_rels/notesMaster1.xml.rels', relsXml([{ id: 'rId1', type: `${REL}/theme`, target: '../theme/theme2.xml' }])])

  // ---- slides
  const out = [], info = []
  for (const s of slides) {
    const variants = mode === 'states' ? s.perState.map((_, c) => c) : [null]
    for (const simClick of variants) {
      const n = out.length + 1
      const r = slideXml(s, simClick, { profile, defaults, deck, media, n, font })
      out.push({ n, slide: s, simClick, ...r })
      if (simClick === null || simClick === s.perState.length - 1) info.push({ no: s.no, ...r.info })
    }
  }
  for (const o of out) {
    files.push([`ppt/slides/slide${o.n}.xml`, o.xml])
    files.push([`ppt/slides/_rels/slide${o.n}.xml.rels`, relsXml([{ id: 'rId1', type: `${REL}/slideLayout`, target: `../slideLayouts/slideLayout${kinds.indexOf(o.slide.kind) + 1}.xml` }, { id: 'rId2', type: `${REL}/notesSlide`, target: `../notesSlides/notesSlide${o.n}.xml` }, ...o.rels])])
    files.push([`ppt/notesSlides/notesSlide${o.n}.xml`, notesSlideXml(mode === 'deck' ? o.slide.note : `${o.slide.no} · click ${o.simClick}`)])
    files.push([`ppt/notesSlides/_rels/notesSlide${o.n}.xml.rels`, relsXml([{ id: 'rId1', type: `${REL}/notesMaster`, target: '../notesMasters/notesMaster1.xml' }, { id: 'rId2', type: `${REL}/slide`, target: `../slides/slide${o.n}.xml` }])])
  }

  // ---- presentation
  const sldIds = out.map((o, i) => 256 + i)
  const sections = mode === 'deck' ? sectionList(out, sldIds) : ''
  const presRels = [
    { id: 'rId1', type: `${REL}/slideMaster`, target: 'slideMasters/slideMaster1.xml' },
    ...out.map((o, i) => ({ id: `rId${i + 2}`, type: `${REL}/slide`, target: `slides/slide${o.n}.xml` })),
  ]
  const nm = `rId${presRels.length + 1}`
  presRels.push({ id: nm, type: `${REL}/notesMaster`, target: 'notesMasters/notesMaster1.xml' },
    { id: `rId${presRels.length + 2}`, type: `${REL}/theme`, target: 'theme/theme1.xml' },
    { id: `rId${presRels.length + 3}`, type: `${REL}/presProps`, target: 'presProps.xml' },
    { id: `rId${presRels.length + 4}`, type: `${REL}/viewProps`, target: 'viewProps.xml' },
    { id: `rId${presRels.length + 5}`, type: `${REL}/tableStyles`, target: 'tableStyles.xml' })
  files.push(['ppt/presentation.xml', `${XMLH}<p:presentation ${X.NSDECL} saveSubsetFonts="1">`
    + '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'
    + `<p:notesMasterIdLst><p:notesMasterId r:id="${nm}"/></p:notesMasterIdLst>`
    + `<p:sldIdLst>${out.map((o, i) => `<p:sldId id="${sldIds[i]}" r:id="rId${i + 2}"/>`).join('')}</p:sldIdLst>`
    + `<p:sldSz cx="${X.E(canvas.w)}" cy="${X.E(canvas.h)}"/><p:notesSz cx="6858000" cy="9144000"/>`
    + '<p:defaultTextStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr></p:defaultTextStyle>'
    + sections + '</p:presentation>'])
  files.push(['ppt/_rels/presentation.xml.rels', relsXml(presRels)])
  files.push(['ppt/presProps.xml', `${XMLH}<p:presentationPr ${X.NSDECL}/>`])
  files.push(['ppt/viewProps.xml', `${XMLH}<p:viewPr ${X.NSDECL}><p:normalViewPr><p:restoredLeft sz="15620"/><p:restoredTop sz="94660"/></p:normalViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`])
  files.push(['ppt/tableStyles.xml', `${XMLH}<a:tblStyleLst xmlns:a="${X.NS.a}" def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>`])
  const now = new Date().toISOString().replace(/\.\d+Z$/, 'Z')
  files.push(['docProps/core.xml', `${XMLH}<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">`
    + `<dc:title>${X.esc(deck.title || '')}</dc:title>${deck.author ? `<dc:creator>${X.esc(deck.author)}</dc:creator><cp:lastModifiedBy>${X.esc(deck.author)}</cp:lastModifiedBy>` : ''}`
    + `<dcterms:created xsi:type="dcterms:W3CDTF">${now}</dcterms:created><dcterms:modified xsi:type="dcterms:W3CDTF">${now}</dcterms:modified></cp:coreProperties>`])
  files.push(['docProps/app.xml', `${XMLH}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"><PresentationFormat>Custom</PresentationFormat><Slides>${out.length}</Slides><Notes>${out.length}</Notes><Application>slidev-addon-authoring</Application></Properties>`])
  files.push(...media.files)
  files.push(['_rels/.rels', relsXml([
    { id: 'rId1', type: `${REL}/officeDocument`, target: 'ppt/presentation.xml' },
    { id: 'rId2', type: 'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties', target: 'docProps/core.xml' },
    { id: 'rId3', type: `${REL}/extended-properties`, target: 'docProps/app.xml' },
  ])])
  const exts = [...new Set(media.files.map(([n]) => n.split('.').pop()))]
  const mime = { png: 'image/png', jpeg: 'image/jpeg', jpg: 'image/jpeg', gif: 'image/gif' }
  const over = (part, type) => `<Override PartName="/${part}" ContentType="${type}"/>`
  const types = `${XMLH}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">`
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/>'
    + exts.map(e => `<Default Extension="${e}" ContentType="${mime[e]}"/>`).join('')
    + over('ppt/presentation.xml', `${CT}.presentation.main+xml`) + over('ppt/slideMasters/slideMaster1.xml', `${CT}.slideMaster+xml`)
    + kinds.map((k, i) => over(`ppt/slideLayouts/slideLayout${i + 1}.xml`, `${CT}.slideLayout+xml`)).join('')
    + out.map(o => over(`ppt/slides/slide${o.n}.xml`, `${CT}.slide+xml`) + over(`ppt/notesSlides/notesSlide${o.n}.xml`, `${CT}.notesSlide+xml`)).join('')
    + over('ppt/notesMasters/notesMaster1.xml', `${CT}.notesMaster+xml`)
    + over('ppt/theme/theme1.xml', 'application/vnd.openxmlformats-officedocument.theme+xml') + over('ppt/theme/theme2.xml', 'application/vnd.openxmlformats-officedocument.theme+xml')
    + over('ppt/presProps.xml', `${CT}.presProps+xml`) + over('ppt/viewProps.xml', `${CT}.viewProps+xml`) + over('ppt/tableStyles.xml', `${CT}.tableStyles+xml`)
    + over('docProps/core.xml', 'application/vnd.openxmlformats-package.core-properties+xml') + over('docProps/app.xml', 'application/vnd.openxmlformats-officedocument.extended-properties+xml')
    + '</Types>'
  return { files: [['[Content_Types].xml', types], ...files], info, slides: out }
}

// PowerPoint sections follow the deck's section slides.
function sectionList(out, sldIds) {
  const groups = []
  out.forEach((o, i) => {
    if (!groups.length || o.slide.layout === 'section') groups.push({ name: o.slide.title || `Slide ${o.slide.no}`, ids: [] })
    groups[groups.length - 1].ids.push(sldIds[i])
  })
  if (!out.some(o => o.slide.layout === 'section')) return ''
  const body = groups.map(g => `<p14:section name="${X.esc(g.name)}" id="{${randomUUID().toUpperCase()}}"><p14:sldIdLst>${g.ids.map(id => `<p14:sldId id="${id}"/>`).join('')}</p14:sldIdLst></p14:section>`).join('')
  return `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">${body}</p14:sectionLst></p:ext></p:extLst>`
}

function slideXml(s, simClick, { profile, defaults, deck, media, n, font }) {
  const rels = []
  const links = new Map()
  const rel = (type, target, external = false) => { const id = `rId${rels.length + 3}`; rels.push({ id, type: `${REL}/${type}`, target, external }); return id }
  const ctx = new X.Ctx({
    themeColors: profile.themeColors, baseFont: font,
    linkRel: url => {
      if (!/^(https?:|mailto:)/i.test(url || '')) return null
      if (!links.has(url)) links.set(url, rel('hyperlink', url, true))
      return links.get(url)
    },
  })
  const imageRel = src => {
    const img = src.startsWith('snapshot:') ? s.snapshots?.get(src.slice(9)) : deck.images.get(src)
    return img ? rel('image', media.add(img)) : null
  }
  const pageBg = hexc(s.states.at(-1).pageBg) || 'FFFFFF'
  const bg = profile.background(s.kind) || pageBg
  let { plan, copies, N } = planSlide(s.perState, bg)
  if (simClick !== null) plan = plan.map(p => ({ ...p, spec: simulate(p, simClick), appear: null, disappear: null, effects: {} })).filter(p => p.spec)
  const kindDefaults = defaults[s.kind] || {}
  let id = 2
  const xml = [], kinds = new Map(), names = new Map()
  const clicks = simClick === null ? Array.from({ length: N }, () => []) : []
  let filler = null
  for (const p of plan) {
    let spec = p.spec
    if (spec.type === 'ph' && !kindDefaults[spec.role]) spec = phAsTextBox(spec)
    const x = emit(spec, id, { ctx, imageRel, deck, kindDefaults, explicit: profile.explicitPlaceholders })
    if (!x) continue
    xml.push(x)
    names.set(p.key, `${spec.name || (spec.type === 'ph' ? X.PH[spec.role][1] : 'Shape')} ${id}`)
    kinds.set(id, { shape: 'sp', ph: 'sp', cxn: 'cxn', pic: 'pic', table: 'gf' }[spec.type])
    if (!filler && p.appear == null && p.disappear == null) {
      // a click with nothing to animate still needs one effect: recolour a single-coloured text to its own colour
      const runs = (spec.paras || spec.text?.paras || []).flatMap(q => q.runs).filter(r => r.t !== undefined)
      const colors = new Set(runs.map(r => r.color || kindDefaults[spec.role]?.color))
      if (runs.length && colors.size === 1 && [...colors][0]) filler = { spid: id, color: [...colors][0], title: spec.type === 'ph' }
    }
    if (p.appear != null) clicks[p.appear - 1].push({ spid: id, kind: 'appear' })
    if (p.disappear != null) clicks[p.disappear - 1].push({ spid: id, kind: 'disappear' })
    for (const [k, evs] of Object.entries(p.effects)) for (const [kind, value, para] of evs) {
      const e = kind === 'bold' ? { spid: id, kind, bold: value } : { spid: id, kind, color: value }
      if (para !== undefined) e.para = para
      clicks[k - 1].push(e)
    }
    id++
  }
  if (profile.slideChrome(s.kind)) { xml.push(X.slideNumberXml(id, s.no), X.footerXml(id + 1, deck.footer)); id += 2 }
  const empty = clicks.map((c, i) => c.length ? 0 : i + 1).filter(Boolean)
  if (filler) for (const i of empty) clicks[i - 1].push({ spid: filler.spid, kind: 'font', color: filler.color })
  const timing = new X.Timing().xml(clicks, kinds)
  const bgXml = !profile.background(s.kind) && pageBg !== 'FFFFFF' ? `<p:bg><p:bgPr><a:solidFill><a:srgbClr val="${pageBg}"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>` : ''
  return {
    xml: `${XMLH}<p:sld ${X.NSDECL}><p:cSld>${bgXml}<p:spTree>${GRP}${xml.join('')}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>${timing}</p:sld>`,
    rels, info: { clicks: N, copies, names, lostClicks: filler ? [] : empty },
  }
}

function phAsTextBox(o) {
  return { type: 'shape', txBox: true, x: o.x, y: o.y, w: o.w + 4, h: o.h, geom: { prst: 'rect' }, fill: null, line: null, order: o.order, name: 'TextBox',
    text: { paras: o.paras, ins: [0, 0, 0, 0], anchor: 't', wrap: 'square', autofit: 'sp' } }
}

function emit(o, id, { ctx, imageRel, deck, kindDefaults, explicit }) {
  if (o.type === 'shape') {
    let x = X.shapeXml(o, id, ctx)
    if (X.hasMath(o)) {
      let fb = null
      if (X.mathOnly(o)) {
        const tex = o.text.paras.flatMap(p => p.runs).find(r => r.omml !== undefined)?.tex || ''
        const eq = deck.equations.get(texKey(tex))
        if (eq) {
          const rid = imageRel('eq:' + texKey(tex))
          const w = eq.w, h = eq.h
          if (rid) fb = X.pictureShapeXml({ ...o, x: o.x + (o.w - w) / 2, y: o.y + (o.h - h) / 2, w, h }, id, rid)
        }
      }
      x = X.wrapMath(x, fb || X.shapeXml(o, id, ctx, true))
    }
    return x
  }
  if (o.type === 'cxn') return X.cxnXml(o, id, ctx)
  if (o.type === 'pic') { const rid = imageRel(o.src); return rid ? X.picXml(o, id, rid) : null }
  if (o.type === 'table') return X.tableXml(o, id, ctx)
  if (o.type === 'ph') {
    const d = kindDefaults[o.role]
    let x = X.phXml(o, id, ctx, d, { explicit })
    if (X.hasMath(o)) x = X.wrapMath(x, X.phXml(o, id, ctx, d, { explicit, fallback: true }))
    return x
  }
  throw new Error('unknown object ' + o.type)
}
