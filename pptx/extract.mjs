// Drive the running deck in Chromium: every slide, every click state.
import { mml2omml } from 'mathml2omml'
import { prepareFonts, extractSlide, fixOmml, rasterize } from './page.mjs'
import { snap, hexc } from './scene.mjs'
import { pts100 } from './xml.mjs'
import { texKey } from './build.mjs'

export async function deckInfo(page, base) {
  await open(page, `${base}/1?embedded=true`, 1)
  return page.evaluate(() => {
    const s = window.__slidev__, c = s.configs
    return {
      canvasWidth: c.canvasWidth, aspectRatio: c.aspectRatio, theme: c.theme, title: c.title, author: c.author || '',
      pptx: c.pptx || {}, footer: c.defaults?.footer,
      slides: s.nav.slides.map(r => ({ no: r.no, layout: r.meta?.slide?.frontmatter?.layout || r.meta?.layout || (r.no === 1 ? 'cover' : 'default'), title: r.meta?.slide?.title || '', note: r.meta?.slide?.note || '' })),
    }
  })
}

async function open(page, url, no) {
  await page.goto(url, { waitUntil: 'load' })
  await page.waitForSelector(`.slidev-page-${no} .slidev-layout`, { timeout: 60000 })
  await page.evaluate(() => document.fonts.ready)
  await page.evaluate(() => Promise.all([...document.images].map(i => i.complete ? 0 : i.decode().catch(() => 0))))
  await page.waitForTimeout(250)
}

const shot = async (page, r, pad = 0, omitBackground = false) => {
  const vp = page.viewportSize()
  const x = Math.max(0, r.x - pad), y = Math.max(0, r.y - pad)
  const width = Math.min(vp.width - x, r.w + 2 * pad), height = Math.min(vp.height - y, r.h + 2 * pad)
  if (width < 1 || height < 1) return null
  return page.screenshot({ clip: { x, y, width, height }, omitBackground })
}

// Returns the slides with their click states, plus the images and equation pictures they use.
export async function extractSlides(page, { base, info, profile, trail, only, keepShots, log }) {
  const canvas = { w: info.canvasWidth, h: Math.round(info.canvasWidth / info.aspectRatio) }
  const images = new Map(), equations = new Map(), fonts = {}
  const slides = []
  for (const meta of info.slides) {
    if (only && !only.has(meta.no)) continue
    const no = meta.no
    const slide = { ...meta, states: [], issues: new Map(), snapshots: new Map(), shots: [], mathRects: [] }
    let clicks = 0
    for (let c = 0; c <= clicks; c++) {
      await open(page, `${base}/${no}?clicks=${c}&embedded=true`, no)
      if (c === 0) clicks = await page.evaluate(() => { const t = window.__slidev__.nav.clicksTotal; return typeof t === 'object' ? t.value : t }) || 0
      for (const [k, v] of Object.entries(await page.evaluate(prepareFonts, { no }))) fonts[k] = (fonts[k] || 0) + v
      await page.evaluate(() => document.fonts.ready)
      await page.waitForTimeout(100)
      const origin = await page.evaluate(no => { const r = document.querySelector(`.slidev-page-${no}`).getBoundingClientRect(); return { x: r.left, y: r.top, s: r.width } }, no)
      const scale = origin.s / canvas.w
      const toPage = r => ({ x: origin.x + r.x * scale, y: origin.y + r.y * scale, w: r.w * scale, h: r.h * scale })
      const st = await page.evaluate(extractSlide, { no, canvasWidth: canvas.w, skip: profile.skip, roles: profile.roles(trail) })

      // equations → OMML, plus a picture for viewers without Office math
      const rects = []
      for (const it of st.items) {
        const paras = it.kind === 'para' ? [it] : it.kind === 'table' ? it.rows.flatMap(r => r.cells.flatMap(c => c.paras)) : []
        for (const p of paras) for (const r of p.runs) {
          if (r.math === undefined) continue
          rects.push(r.rect)
          const key = texKey(r.tex)
          try {
            const mathml = r.math.replace(/<annotation[\s\S]*?<\/annotation>/g, '')
            r.omml = await page.evaluate(fixOmml, { omml: mml2omml(mathml), sz: pts100(snap(r.size)), color: hexc(r.color), display: r.display })
          } catch {
            r.omml = ''
            st.issues.push({ kind: 'math', id: r.mathId, rect: r.rect, label: r.tex.slice(0, 60), detail: '수식을 PowerPoint 수식으로 바꾸지 못해 글자(TeX)로 넣었습니다.', action: 'TeX 글자로 대체', severity: 'fix' })
          }
          if (!equations.has(key)) {
            const png = await shot(page, toPage(r.rect), 0, true)
            if (png) { equations.set(key, { w: r.rect.w, h: r.rect.h }); images.set('eq:' + key, { data: png, ext: 'png' }) }
          }
        }
      }
      slide.mathRects.push(rects)
      for (const it of st.items) if (it.kind === 'snapshot' && !slide.snapshots.has(it.id)) {
        const png = await shot(page, toPage(it.rect))
        if (png) slide.snapshots.set(it.id, { data: png, ext: 'png' })
      }
      for (const is of st.issues) {
        const key = `${is.kind}|${is.id}`
        if (slide.issues.has(key)) { slide.issues.get(key).clicks.push(c); continue }
        slide.issues.set(key, { ...is, clicks: [c], crop: await shot(page, toPage(is.rect), 16) })
      }
      if (keepShots) slide.shots.push(await shot(page, toPage({ x: 0, y: 0, w: canvas.w, h: canvas.h })))
      slide.states.push({ pageBg: st.pageBg, items: st.items })
    }
    slides.push(slide)
    log?.(no, clicks)
  }
  return { canvas, slides, images, equations, fonts }
}

const kindOf = b => b[0] === 0x89 && b[1] === 0x50 ? 'png' : b[0] === 0xff && b[1] === 0xd8 ? 'jpeg' : b[0] === 0x47 && b[1] === 0x49 ? 'gif' : null

// Every picture the slides reference, as bytes PowerPoint accepts.
export async function fetchImages(page, slides, images) {
  const srcs = new Set()
  for (const s of slides) for (const st of s.states) for (const it of st.items) {
    if (it.kind === 'image') srcs.add(it.src)
    if (it.kind === 'svg' && it.tag === 'image') srcs.add(it.geom.href)
  }
  const failed = [], rasterized = []
  for (const src of srcs) {
    if (images.has(src)) continue
    try {
      let data
      if (src.startsWith('data:')) data = Buffer.from(src.split(',')[1], /;base64/.test(src.split(',')[0]) ? 'base64' : 'utf8')
      else { const res = await page.request.get(src); if (!res.ok()) throw new Error(res.status()); data = await res.body() }
      let ext = kindOf(data)
      if (!ext) {
        const url = await page.evaluate(rasterize, { url: src, w: 0, h: 0 })
        data = Buffer.from(url.split(',')[1], 'base64'); ext = 'png'
        rasterized.push(src)
      }
      images.set(src, { data, ext })
    } catch (e) { failed.push(`${src} (${e.message})`) }
  }
  return { failed, rasterized }
}
