// slidev-pptx: convert a Slidev deck into an editable PowerPoint file with a report of what needs manual work.
import { parseArgs } from 'node:util'
import { spawn } from 'node:child_process'
import { createRequire } from 'node:module'
import { createServer } from 'node:net'
import { existsSync, readFileSync } from 'node:fs'
import { mkdtemp, mkdir, readFile, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { basename, dirname, join, resolve } from 'node:path'
import { pathToFileURL } from 'node:url'
import { deckInfo, extractSlides, fetchImages } from './extract.mjs'
import { buildPackage } from './build.mjs'
import { hcs, generic } from './profiles.mjs'
import { zip } from './zip.mjs'
import { xmlErrors } from './page.mjs'
import { findRenderer, renderSlides, compareStates } from './check.mjs'
import { writeReport, collectNotes } from './report.mjs'

const HELP = `사용법: slidev-pptx [slides.md] [옵션]

  -o, --out <file>        PPTX 경로 (기본: <덱 폴더 이름>.pptx)
      --slides <범위>      일부만 변환, 예: 1-5,8
      --url <url>         이미 실행 중인 Slidev 서버를 사용
      --trail <selector>  레이아웃의 머리글 자리에 넣을 요소 (HCS 테마)
      --profile <이름>     hcs | generic (기본: 테마에 따라 자동)
      --no-check          LibreOffice 모양 비교를 하지 않음
      --threshold <퍼센트> 모양 비교에서 의심으로 표시할 차이 (기본: 3)
      --strict            직접 고칠 항목이 있으면 종료 코드 2
  -h, --help`

const log = msg => process.stderr.write(msg + '\n')

function parseRange(text) {
  if (!text) return null
  const out = new Set()
  for (const part of text.split(',')) {
    const [a, b] = part.split('-').map(Number)
    for (let i = a; i <= (b || a); i++) out.add(i)
  }
  return out
}

const freePort = () => new Promise((res, rej) => { const s = createServer(); s.listen(0, '127.0.0.1', () => { const { port } = s.address(); s.close(() => res(port)) }); s.on('error', rej) })

async function startSlidev(entry, require) {
  const pkgPath = require.resolve('@slidev/cli/package.json')
  const pkg = JSON.parse(readFileSync(pkgPath, 'utf8'))
  const bin = join(dirname(pkgPath), typeof pkg.bin === 'string' ? pkg.bin : pkg.bin.slidev)
  const port = await freePort()
  const child = spawn(process.execPath, [bin, entry, '--port', String(port), '--bind', '127.0.0.1'], { stdio: ['ignore', 'pipe', 'pipe'], detached: process.platform !== 'win32' })
  let output = '', exited = null
  child.stdout.on('data', d => { output += d })
  child.stderr.on('data', d => { output += d })
  child.on('exit', code => { exited = code })
  const url = `http://127.0.0.1:${port}`
  for (let i = 0; i < 600; i++) {
    if (exited !== null) throw new Error(`Slidev가 시작되지 않았습니다:\n${output.slice(-2000)}`)
    try { if ((await fetch(url)).ok) break } catch {}
    await new Promise(r => setTimeout(r, 200))
  }
  const stop = () => { try { process.kill(-child.pid) } catch { child.kill() } }
  return { url, stop }
}

async function loadChromium(require) {
  for (const req of [require, createRequire(import.meta.url)]) {
    try {
      const mod = await import(pathToFileURL(req.resolve('playwright-chromium')).href)
      return mod.chromium ?? mod.default.chromium
    } catch {}
  }
  throw new Error('playwright-chromium이 필요합니다: npm install -D playwright-chromium && npx playwright install chromium')
}

function themeImages(profile, require) {
  const out = {}
  if (!profile.images.length) return out
  try {
    const dir = join(dirname(require.resolve('slidev-theme-hcs/package.json')), 'assets')
    for (const name of profile.images) if (existsSync(join(dir, name + '.png'))) out[name] = { data: readFileSync(join(dir, name + '.png')), ext: 'png' }
  } catch {}
  return out
}

export async function main(argv) {
  const { values, positionals } = parseArgs({
    args: argv, allowPositionals: true, allowNegative: true,
    options: {
      out: { type: 'string', short: 'o' }, slides: { type: 'string' }, url: { type: 'string' }, trail: { type: 'string' },
      profile: { type: 'string' }, check: { type: 'boolean', default: true }, threshold: { type: 'string', default: '3' },
      strict: { type: 'boolean', default: false }, help: { type: 'boolean', short: 'h' },
    },
  })
  if (values.help) { console.log(HELP); return 0 }
  const entry = resolve(positionals[0] || 'slides.md')
  if (!existsSync(entry)) throw new Error(`슬라이드 파일이 없습니다: ${entry}`)
  const deckDir = dirname(entry)
  const out = resolve(values.out || `${basename(deckDir)}.pptx`)
  const reportPath = out.replace(/\.pptx$/i, '') + '.report.md'
  const require = createRequire(join(deckDir, 'noop.js'))
  const chromium = await loadChromium(require)

  const server = values.url ? { url: values.url.replace(/\/$/, ''), stop() {} } : await startSlidev(entry, require)
  const tmp = await mkdtemp(join(tmpdir(), 'slidev-pptx-'))
  const browser = await chromium.launch()
  try {
    const page = await browser.newPage({ viewport: { width: 960, height: 540 }, deviceScaleFactor: 2 })
    const info = await deckInfo(page, server.url)
    const profile = { hcs, generic }[values.profile] || (/hcs/.test(info.theme || '') ? hcs : generic)
    const canvas = { w: info.canvasWidth, h: Math.round(info.canvasWidth / info.aspectRatio) }
    await page.setViewportSize({ width: canvas.w, height: canvas.h })
    const only = parseRange(values.slides)
    const total = only ? only.size : info.slides.length
    let done = 0
    const renderer = values.check ? await findRenderer() : { kind: null, reason: '--no-check로 건너뛰었습니다.' }
    const extracted = await extractSlides(page, {
      base: server.url, info, profile, trail: values.trail || info.pptx.trail, only, keepShots: !!renderer.kind,
      log: (no, clicks) => process.stderr.write(`\r슬라이드 읽는 중 ${++done}/${total} (${no}쪽, 클릭 ${clicks})   `),
    })
    process.stderr.write('\n')
    const { failed, rasterized } = await fetchImages(page, extracted.slides, extracted.images)
    const deck = {
      title: info.title, author: info.author, canvas, profile, footer: info.footer ?? 'HCS Lab @ SNU',
      themeImages: themeImages(profile, require), slides: extracted.slides, images: extracted.images, equations: extracted.equations,
    }
    const pkg = buildPackage(deck)
    const xmlParts = pkg.files.filter(([n]) => /\.(xml|rels)$/.test(n))
    const errors = await page.evaluate(xmlErrors, xmlParts)
    if (errors.length) throw new Error('생성한 XML이 올바르지 않습니다:\n' + errors.join('\n'))
    await mkdir(dirname(out), { recursive: true })
    await writeFile(out, zip(pkg.files))
    log(`PPTX: ${out}`)

    // names of shapes the report points to
    const infoByNo = new Map(pkg.info.map(i => [i.no, i]))
    for (const s of extracted.slides) for (const is of s.issues.values()) if (is.kind === 'snapshot') is.shape = infoByNo.get(s.no)?.names.get('snap:' + is.id)

    let check = { ran: false, reason: renderer.reason, flagged: [], threshold: Number(values.threshold) }
    if (renderer.kind) {
      try {
        log('모양 비교: LibreOffice로 클릭 단계마다 그리는 중…')
        const states = buildPackage(deck, { mode: 'states' })
        const statesPath = join(tmp, 'states.pptx')
        await writeFile(statesPath, zip(states.files))
        const renders = await renderSlides(statesPath, tmp, renderer, log)
        if (process.env.SLIDEV_PPTX_DEBUG_DIR) {
          // development aid: keep the per-state deck, its renders, and the browser shots
          const dbg = process.env.SLIDEV_PPTX_DEBUG_DIR
          await mkdir(dbg, { recursive: true })
          await writeFile(join(dbg, 'states.pptx'), await readFile(statesPath))
          for (const [i, r] of renders.entries()) await writeFile(join(dbg, `render-${String(i).padStart(3, '0')}.png`), await readFile(r))
          let k = 0
          for (const s of extracted.slides) for (const [c, shot] of s.shots.entries()) await writeFile(join(dbg, `shot-${String(k++).padStart(3, '0')}-s${s.no}-c${c}.png`), shot)
        }
        const list = extracted.slides.flatMap(s => s.shots.map((shot, c) => ({ no: s.no, click: c, shot, masks: s.mathRects[c] || [] })))
        const cmp = await browser.newPage()
        check = { ...check, ran: true, flagged: await compareStates(cmp, list, renders, canvas, check.threshold) }
      } catch (e) {
        check.reason = `LibreOffice로 그리지 못했습니다(${e.message.split('\n')[0]}).`
      }
    }
    const notes = collectNotes({ slides: extracted.slides, info: pkg.info, fonts: extracted.fonts, rasterized })
    const summary = await writeReport(reportPath, {
      title: info.title, out: basename(out), slideCount: extracted.slides.length, clickCount: extracted.slides.reduce((a, s) => a + s.states.length - 1, 0),
      slides: extracted.slides, imageFailures: failed, check, notes,
    })
    log(`보고서: ${reportPath}`)
    log(`직접 고칠 항목 ${summary.fixes}개 · ${summary.flagged === null ? '모양 비교 안 함' : `모양 비교 의심 ${summary.flagged}곳`} · PowerPoint에서 확인할 것 ${summary.notes}가지`)
    return values.strict && summary.fixes ? 2 : 0
  } finally {
    await browser.close()
    server.stop()
    await rm(tmp, { recursive: true, force: true })
  }
}
