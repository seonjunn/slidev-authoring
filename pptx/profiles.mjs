// Slide frames for the generated deck. The HCS profile rebuilds slidev-theme-hcs as a master with
// layouts and placeholders; the generic profile works with any theme and keeps its decoration as
// ordinary shapes on each slide.
import { E, NSDECL, esc } from './xml.mjs'

const GRP = '<p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
const xfrm = (x, y, w, h) => `<a:xfrm><a:off x="${E(x)}" y="${E(y)}"/><a:ext cx="${E(w)}" cy="${E(h)}"/></a:xfrm>`
const pts = px => Math.round(px * 75)

export function themeXml({ name, colors, font }) {
  const c = (tag, v) => `<a:${tag}><a:srgbClr val="${v}"/></a:${tag}>`
  return `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${esc(name)}"><a:themeElements>`
    + `<a:clrScheme name="${esc(name)}">${c('dk1', colors.dk1)}${c('lt1', colors.lt1)}${c('dk2', colors.dk2)}${c('lt2', colors.lt2)}`
    + ['accent1', 'accent2', 'accent3', 'accent4', 'accent5', 'accent6', 'hlink', 'folHlink'].map(k => c(k, colors[k])).join('') + '</a:clrScheme>'
    + `<a:fontScheme name="${esc(name)}">`
    + `<a:majorFont><a:latin typeface="${esc(font)}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Jpan" typeface="游ゴシック"/></a:majorFont>`
    + `<a:minorFont><a:latin typeface="${esc(font)}"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Jpan" typeface="游ゴシック"/></a:minorFont></a:fontScheme>`
    + '<a:fmtScheme name="Office">'
    + '<a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="50000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="80000"/></a:schemeClr></a:solidFill></a:fillStyleLst>'
    + '<a:lnStyleLst>' + [6350, 12700, 19050].map(w => `<a:ln w="${w}" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln>`).join('') + '</a:lnStyleLst>'
    + '<a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle></a:effectStyleLst>'
    + '<a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/></a:schemeClr></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:shade val="90000"/></a:schemeClr></a:solidFill></a:bgFillStyleLst>'
    + '</a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/></a:theme>'
}

function ph(id, name, phx, box, body = '<a:bodyPr/>', paras = '', lst = '') {
  return `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="${name}"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr>${phx}</p:nvPr></p:nvSpPr>`
    + `<p:spPr>${xfrm(...box)}</p:spPr><p:txBody>${body}<a:lstStyle>${lst}</a:lstStyle>${paras}</p:txBody></p:sp>`
}
const prompt = (text, lvl) => `<a:p>${lvl ? `<a:pPr lvl="${lvl}"/>` : ''}<a:r><a:rPr lang="en-US"/><a:t>${text}</a:t></a:r></a:p>`
const bodyPrompts = () => ['Click to edit Master text styles', 'Second level', 'Third level'].map((t, i) => prompt(t, i)).join('')
const ZERO = (anchor, fit = '<a:normAutofit/>') => `<a:bodyPr vert="horz" lIns="0" tIns="0" rIns="0" bIns="0" rtlCol="0" anchor="${anchor}">${fit}</a:bodyPr>`
const defRPr = (d, extra = '') => `<a:defRPr sz="${pts(d.sz)}"${d.b !== undefined ? ` b="${d.b ? 1 : 0}"` : ''}${extra}>${d.color ? `<a:solidFill><a:srgbClr val="${d.color}"/></a:solidFill>` : ''}</a:defRPr>`

function lvl(i, marL, sz, char, font = 'Arial') {
  return `<a:lvl${i}pPr marL="${E(marL)}" indent="${-E(i === 1 ? 24 : 22)}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">`
    + '<a:lnSpc><a:spcPct val="122000"/></a:lnSpc><a:spcBef><a:spcPts val="0"/></a:spcBef><a:spcAft><a:spcPts val="600"/></a:spcAft>'
    + `<a:buFont typeface="${font}"/><a:buChar char="${char}"/><a:defRPr sz="${sz}" kern="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl${i}pPr>`
}

function masterXml({ layoutIds, title, titleBox, bodyBox, extra = '', bg = '<p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg>' }) {
  let sp = ph(2, 'Title Placeholder 1', '<p:ph type="title"/>', titleBox, ZERO('t'), prompt('Click to edit Master title style'))
  sp += ph(3, 'Text Placeholder 2', '<p:ph type="body" idx="1"/>', bodyBox, ZERO('t'), bodyPrompts())
  sp += extra
  const titleStyle = '<p:titleStyle><a:lvl1pPr algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:lnSpc><a:spcPct val="100000"/></a:lnSpc><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/>'
    + `<a:defRPr sz="${pts(title.sz)}" b="${title.b ? 1 : 0}" kern="0"><a:solidFill><a:srgbClr val="${title.color}"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr></p:titleStyle>`
  const bodyStyle = '<p:bodyStyle>' + lvl(1, 24, 1500, '•') + lvl(2, 46, 1400, 'o', 'Courier New') + lvl(3, 68, 1200, '▪') + '</p:bodyStyle>'
  const other = '<p:otherStyle><a:defPPr><a:defRPr lang="en-US"/></a:defPPr><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">'
    + '<a:defRPr sz="1800" kern="0"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:otherStyle>'
  return `${XML}<p:sldMaster ${NSDECL}><p:cSld>${bg}<p:spTree>${GRP}${sp}</p:spTree></p:cSld>`
    + '<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
    + `<p:sldLayoutIdLst>${layoutIds}</p:sldLayoutIdLst><p:txStyles>${titleStyle}${bodyStyle}${other}</p:txStyles></p:sldMaster>`
}

const layoutXml = ({ name, type, bg = '', shapes }) => `${XML}<p:sldLayout ${NSDECL}${type ? ` type="${type}"` : ''} preserve="1"><p:cSld name="${name}">${bg}<p:spTree>${GRP}${shapes}</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sldLayout>`

function pic(id, name, rid, box, cropRight) {
  return `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${esc(name)}" descr="${esc(name)}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr userDrawn="1"/></p:nvPicPr>`
    + `<p:blipFill><a:blip r:embed="${rid}"/>${cropRight ? `<a:srcRect r="${cropRight}"/>` : ''}<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>${xfrm(...box)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`
}

// Most common value of fn over the placeholders of one role, used as the layout default.
export function modeOf(values) {
  const n = new Map()
  for (const v of values) { const k = JSON.stringify(v); n.set(k, (n.get(k) || 0) + 1) }
  if (!n.size) return null
  return JSON.parse([...n.entries()].reduce((a, b) => b[1] > a[1] ? b : a)[0])
}

// ------------------------------------------------------------------ HCS ---
const HCS = {
  navy: '2D3250', heading: '424769', body: '434343',
  colors: { dk1: '434343', lt1: 'FFFFFF', dk2: '2D3250', lt2: 'F2F2F2', accent1: '424769', accent2: '7077A1', accent3: '8E75AD', accent4: '7DA294', accent5: 'F6B17A', accent6: 'A7B8CA', hlink: '555F86', folHlink: '8E75AD' },
}
const BAR = 540 - 18.08

export const hcs = {
  name: 'hcs',
  skip: ['.hcs-footer', '.hcs-brand', '.hcs-cover-rule', '.figure-edit-button'],
  roles: trail => [
    { role: 'ctrTitle', match: 'h1', within: '.hcs-cover-title' },
    { role: 'subTitle', match: '.hcs-cover-author, .hcs-cover-subtitle' },
    { role: 'title', match: 'h1', within: '.hcs-center' },
    { role: 'title', match: 'h1', within: '.hcs-content', first: true },
    ...(trail ? [{ role: 'trail', match: trail }] : []),
  ],
  kinds: ['cover', 'section', 'content', 'end'],
  kindOf: layout => ({ cover: 'cover', section: 'section', end: 'end' })[layout] || 'content',
  themeColors: { '434343': 'tx1', FFFFFF: 'bg1', '2D3250': 'tx2', F2F2F2: 'bg2', '424769': 'accent1', '7077A1': 'accent2', '8E75AD': 'accent3', '7DA294': 'accent4', F6B17A: 'accent5' },
  background: kind => ['section', 'end'].includes(kind) ? HCS.navy : 'FFFFFF',
  // Placeholder defaults come from what the deck actually shows, so deck CSS overrides carry over.
  defaults(observed) {
    const pick = (kind, role, fallback) => ({ ...fallback, ...(observed(kind, role) || {}) })
    return {
      content: { title: pick('content', 'title', { sz: 32, b: true, color: HCS.heading, font: 'Arial', algn: 'l' }), trail: pick('content', 'trail', { sz: 10.5, color: '9297A1', font: 'Arial', algn: 'l' }) },
      cover: { ctrTitle: pick('cover', 'ctrTitle', { sz: 53.333, color: HCS.heading, font: 'Arial', algn: 'ctr' }), subTitle: pick('cover', 'subTitle', { sz: 20, color: 'B7B7B7', font: 'Arial', algn: 'ctr' }) },
      section: { title: pick('section', 'title', { sz: 48, color: 'F2F2F2', font: 'Arial', algn: 'ctr' }) },
      end: { title: pick('end', 'title', { sz: 53.333, color: 'F2F2F2', font: 'Arial', algn: 'ctr' }) },
    }
  },
  theme: () => themeXml({ name: 'HCS', colors: HCS.colors, font: 'Arial' }),
  images: ['snu-color', 'hcs-color', 'snu-white-full', 'hcs-white', 'hcs-full-white'],
  master({ layoutIds, defaults, footer }) {
    const t = defaults.content.title
    const foot = '<a:lvl1pPr algn="r"><a:defRPr sz="800"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:lvl1pPr>'
    const extra = ph(4, 'Footer Placeholder 3', '<p:ph type="ftr" sz="quarter" idx="3"/>', [600, BAR, 352, 18.08], ZERO('ctr', '<a:noAutofit/>'), `<a:p><a:r><a:rPr lang="en-US"/><a:t>${esc(footer)}</a:t></a:r></a:p>`, foot)
      + ph(5, 'Slide Number Placeholder 4', '<p:ph type="sldNum" sz="quarter" idx="4"/>', [8, BAR, 120, 18.08], ZERO('ctr', '<a:noAutofit/>'), '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>', foot.replace('algn="r"', 'algn="l"'))
    return masterXml({ layoutIds, title: t, titleBox: [42.32, 25.65, 901.68, 35.4], bodyBox: [42.32, 75, 875.36, 397], extra })
  },
  layout(kind, { defaults, footer, rids }) {
    const navy = '<p:bg><p:bgPr><a:solidFill><a:schemeClr val="tx2"/></a:solidFill><a:effectLst/></p:bgPr></p:bg>'
    const center = [42.34, 225.83, 875.33, 88.34]
    const d = defaults[kind]
    if (kind === 'content') {
      const foot = '<a:lvl1pPr algn="r"><a:defRPr sz="800"><a:solidFill><a:schemeClr val="bg1"/></a:solidFill></a:defRPr></a:lvl1pPr>'
      const shapes = ph(2, 'Title 1', '<p:ph type="title"/>', [42.32, 25.65, 901.68, 35.4], '<a:bodyPr/>', prompt('Click to edit Master title style'))
        + ph(3, 'Header Placeholder 2', '<p:ph type="body" sz="quarter" idx="13"/>', [42.32, 7, 720, 12.6], ZERO('t', '<a:noAutofit/>'), prompt('Section / topic'),
          `<a:lvl1pPr marL="0" indent="0"><a:lnSpc><a:spcPct val="104000"/></a:lnSpc><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/>${defRPr(d.trail)}</a:lvl1pPr>`)
        + `<p:sp><p:nvSpPr><p:cNvPr id="4" name="Footer Bar"/><p:cNvSpPr/><p:nvPr userDrawn="1"/></p:nvSpPr><p:spPr>${xfrm(0, BAR, 960, 18.08)}<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:solidFill><a:schemeClr val="accent1"/></a:solidFill><a:ln><a:noFill/></a:ln></p:spPr><p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp>`
        + ph(5, 'Slide Number Placeholder', '<p:ph type="sldNum" sz="quarter" idx="12"/>', [8, BAR, 120, 18.08], ZERO('ctr', '<a:noAutofit/>'), '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>‹#›</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p>', foot.replace('algn="r"', 'algn="l"'))
        + ph(6, 'Footer Placeholder', '<p:ph type="ftr" sz="quarter" idx="11"/>', [600, BAR, 352, 18.08], ZERO('ctr', '<a:noAutofit/>'), `<a:p><a:r><a:rPr lang="en-US"/><a:t>${esc(footer)}</a:t></a:r><a:endParaRPr lang="en-US"/></a:p>`, foot)
      return layoutXml({ name: 'Title Only', type: 'titleOnly', shapes })
    }
    if (kind === 'cover') {
      let shapes = ph(2, 'Title 1', '<p:ph type="ctrTitle"/>', [20, 85, 920, 190], ZERO('b'), prompt('Click to edit Master title style'), `<a:lvl1pPr algn="ctr"><a:lnSpc><a:spcPct val="96000"/></a:lnSpc>${defRPr({ ...d.ctrTitle, b: d.ctrTitle.b || false })}</a:lvl1pPr>`)
        + ph(3, 'Subtitle 2', '<p:ph type="subTitle" idx="1"/>', [190, 378, 580, 48], ZERO('t'), prompt('Click to edit Master subtitle style'),
          `<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:lnSpc><a:spcPct val="104000"/></a:lnSpc><a:spcAft><a:spcPts val="0"/></a:spcAft><a:buNone/>${defRPr(d.subTitle)}</a:lvl1pPr>`)
        + `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="4" name="Title Rule"/><p:cNvCxnSpPr/><p:nvPr userDrawn="1"/></p:nvCxnSpPr><p:spPr>${xfrm(423.65, 348.9, 112.75, 0)}<a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="${E(1)}"><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:ln></p:spPr></p:cxnSp>`
      if (rids['snu-color']) shapes += pic(5, 'Seoul National University', rids['snu-color'], [326.08, 457.08, 155.43, 66.93])
      if (rids['hcs-color']) shapes += pic(6, 'HCS', rids['hcs-color'], [494, 455.9, 124.67, 69.29])
      return layoutXml({ name: 'Title Slide', type: 'title', shapes })
    }
    const titleLst = `<a:lvl1pPr algn="ctr">${defRPr({ ...d.title, b: d.title.b || false })}</a:lvl1pPr>`
    let shapes = ph(2, 'Title 1', '<p:ph type="title"/>', center, ZERO('ctr'), prompt('Click to edit Master title style'), titleLst)
    if (kind === 'section') {
      if (rids['snu-white-full']) shapes += pic(3, 'Seoul National University', rids['snu-white-full'], [11.22, 7.45, 36, 36], 88070)
      if (rids['hcs-white']) shapes += pic(4, 'HCS', rids['hcs-white'], [47.22, 8.49, 61.04, 33.92])
      return layoutXml({ name: 'Section Header', type: 'secHead', bg: navy, shapes })
    }
    if (rids['snu-white-full']) shapes += pic(3, 'Seoul National University', rids['snu-white-full'], [298.11, 446.45, 74.44, 74.42], 88070)
    if (rids['hcs-full-white']) shapes += pic(4, 'Human-Centered Computer Systems Lab', rids['hcs-full-white'], [374.15, 434.27, 287.8, 85.95])
    return layoutXml({ name: 'Closing', bg: navy, shapes })
  },
  layoutImages: { cover: ['snu-color', 'hcs-color'], section: ['snu-white-full', 'hcs-white'], end: ['snu-white-full', 'hcs-full-white'], content: [] },
  explicitPlaceholders: false,
  slideChrome: kind => kind === 'content',
}

// -------------------------------------------------------------- generic ---
export const generic = {
  name: 'generic',
  skip: ['.figure-edit-button'],
  roles: trail => [{ role: 'title', match: 'h1', first: true }, ...(trail ? [{ role: 'trail', match: trail }] : [])],
  kinds: ['content', 'blank'],
  kindOf: () => 'content',
  themeColors: {},
  background: () => null,
  defaults(observed) {
    const t = observed('content', 'title') || {}
    return { content: { title: { sz: 36, b: true, color: '222222', font: 'Arial', algn: 'l', ...t }, trail: { sz: 10.5, color: '888888', font: 'Arial', algn: 'l', ...(observed('content', 'trail') || {}) } }, blank: {} }
  },
  theme: ({ font }) => themeXml({ name: 'Slidev', font, colors: { dk1: '222222', lt1: 'FFFFFF', dk2: '44546A', lt2: 'E7E6E6', accent1: '4472C4', accent2: 'ED7D31', accent3: 'A5A5A5', accent4: 'FFC000', accent5: '5B9BD5', accent6: '70AD47', hlink: '0563C1', folHlink: '954F72' } }),
  images: [],
  master({ layoutIds, defaults, size, titleBox }) {
    return masterXml({ layoutIds, title: defaults.content.title, titleBox, bodyBox: [titleBox[0], titleBox[1] + titleBox[3] + 16, size.w - 2 * titleBox[0], size.h - titleBox[1] - titleBox[3] - 48] })
  },
  layout(kind, { titleBox }) {
    if (kind === 'blank') return layoutXml({ name: 'Blank', type: 'blank', shapes: '' })
    return layoutXml({ name: 'Title Only', type: 'titleOnly', shapes: ph(2, 'Title 1', '<p:ph type="title"/>', titleBox, '<a:bodyPr/>', prompt('Click to edit Master title style')) })
  },
  layoutImages: { content: [], blank: [] },
  explicitPlaceholders: true,
  slideChrome: () => false,
}

export const notesMasterXml = () => `${XML}<p:notesMaster ${NSDECL}><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>${GRP}`
  + `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg" idx="2"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="1143000"/><a:ext cx="5486400" cy="3086100"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln w="12700"><a:solidFill><a:prstClr val="black"/></a:solidFill></a:ln></p:spPr></p:sp>`
  + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" sz="quarter" idx="3"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="685800" y="4400550"/><a:ext cx="5486400" cy="3600450"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody><a:bodyPr vert="horz" lIns="91440" tIns="45720" rIns="91440" bIns="45720" rtlCol="0"/><a:lstStyle/><a:p><a:pPr lvl="0"/><a:r><a:rPr lang="en-US"/><a:t>Click to edit Master text styles</a:t></a:r></a:p></p:txBody></p:sp>`
  + '</p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
  + '<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle></p:notesMaster>'

// Speaker notes: one paragraph per source paragraph; a Markdown heading becomes a bold paragraph.
export function notesSlideXml(note) {
  const paras = String(note || '').trim().split(/\n\s*\n/).map(p => p.trim()).filter(Boolean)
  const body = paras.map(text => {
    const heading = text.startsWith('#')
    text = text.replace(/^#+\s*/, '').replace(/`([^`]+)`/g, '$1').replace(/\*\*([^*]+)\*\*/g, '$1').replace(/\s*\n\s*/g, ' ')
    return `<a:p><a:r><a:rPr lang="ko-KR" altLang="en-US"${heading ? ' b="1"' : ''} dirty="0"/><a:t>${esc(text)}</a:t></a:r></a:p>`
  }).join('') || '<a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p>'
  return `${XML}<p:notes ${NSDECL}><p:cSld><p:spTree>${GRP}`
    + '<p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp>'
    + `<p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${body}</p:txBody></p:sp>`
    + '</p:spTree></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>'
}
