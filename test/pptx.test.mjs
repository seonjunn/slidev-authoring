import test from 'node:test'
import assert from 'node:assert/strict'
import { inflateRawSync } from 'node:zlib'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { parsePath, bbox, pointsOf } from '../pptx/geom.mjs'
import { zip } from '../pptx/zip.mjs'
import { normRuns, pathKey, cmpKey } from '../pptx/scene.mjs'
import { planKey } from '../pptx/plan.mjs'
import { buildPackage } from '../pptx/build.mjs'
import { hcs, generic } from '../pptx/profiles.mjs'
import { writeReport } from '../pptx/report.mjs'

function unzip(buf) {
  const files = new Map()
  const end = buf.lastIndexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06]))
  let p = buf.readUInt32LE(end + 16)
  for (let i = 0; i < buf.readUInt16LE(end + 10); i++) {
    const method = buf.readUInt16LE(p + 10), size = buf.readUInt32LE(p + 20), nameLen = buf.readUInt16LE(p + 28)
    const extra = buf.readUInt16LE(p + 30), comment = buf.readUInt16LE(p + 32), local = buf.readUInt32LE(p + 42)
    const name = buf.toString('utf8', p + 46, p + 46 + nameLen)
    const start = local + 30 + buf.readUInt16LE(local + 26) + buf.readUInt16LE(local + 28)
    const body = buf.subarray(start, start + size)
    files.set(name, method === 8 ? inflateRawSync(body) : body)
    p += 46 + nameLen + extra + comment
  }
  return files
}

const run = (t, extra = {}) => ({ t, size: 20, weight: 400, italic: false, color: { r: 66, g: 71, b: 105, a: 1 }, font: 'Arial', va: null, u: false, s: false, hl: null, ls: 0, caps: false, ws: 'normal', link: null, ...extra })
const para = (id, text, extra = {}) => ({
  kind: 'para', id, runs: [run(text)], align: 'start', fs: 20, lh: 28, pre: false, li: null, onlyChild: true, whiteSpace: 'normal', boxChain: [], svgGroup: null, role: null,
  lines: [{ top: 100, bottom: 128, left: 50, right: 300, maxh: 28 }], content: { x: 50, y: 100, w: 500, h: 28 }, ...extra,
})

test('path parsing turns arcs and relative commands into absolute segments', () => {
  const subs = parsePath('M10 10 h20 v20 a10 10 0 0 1 -20 0 z')
  assert.equal(subs.length, 1)
  assert.ok(subs[0].closed)
  const [x0, y0, x1, y1] = bbox(pointsOf(subs))
  assert.deepEqual([x0, y0, x1], [10, 10, 30])
  assert.ok(y1 > 30 && y1 < 45)
})

test('zip output reads back to the same entries', () => {
  const big = 'x'.repeat(5000)
  const files = unzip(zip([['[Content_Types].xml', '<Types/>'], ['ppt/a.xml', big], ['한글.txt', '가나다']]))
  assert.deepEqual([...files.keys()], ['[Content_Types].xml', 'ppt/a.xml', '한글.txt'])
  assert.equal(files.get('ppt/a.xml').toString(), big)
  assert.equal(files.get('한글.txt').toString(), '가나다')
})

test('runs collapse whitespace across boundaries and merge equal formatting', () => {
  const runs = normRuns({ pre: false, whiteSpace: 'normal', fs: 20, runs: [run('  Hello \n'), run(' world '), run('bold', { weight: 700 })] })
  assert.deepEqual(runs.map(r => r.t), ['Hello world ', 'bold'])
  assert.equal(runs[1].b, true)
})

test('DOM path keys sort in document order', () => {
  const ids = ['P/div0/p2', 'P/div0/p10', 'P/div0', 'P/div0/p2#3']
  assert.deepEqual([...ids].sort((a, b) => cmpKey(pathKey(a), pathKey(b))), ['P/div0', 'P/div0/p2', 'P/div0/p2#3', 'P/div0/p10'])
})

const box = (fill, line = { color: 'AAA0BE', w: 1, dash: null, alpha: 1 }, text = 'A') => ({
  type: 'shape', x: 10, y: 10, w: 100, h: 40, geom: { prst: 'rect' }, fill: fill && { color: fill, alpha: 1 }, line, order: [0, 1], name: 'Rectangle',
  text: { paras: [{ runs: [{ t: text, sz: 20, color: '424769', font: 'Arial' }], algn: 'l', lnSpc: ['pct', 100000] }], ins: [0, 0, 0, 0], anchor: 't', wrap: 'square', autofit: 'none' },
})

test('a colour-only highlight stays one shape with emphasis effects', () => {
  const plan = planKey([box('F0EDF6'), box('E3DAEE'), box('F0EDF6')], 2, 'FFFFFF')
  assert.equal(plan.length, 1)
  assert.deepEqual(plan[0].effects, { 1: [['fill', 'E3DAEE']], 2: [['fill', 'F0EDF6']] })
})

test('a text change becomes copies that appear and disappear', () => {
  const plan = planKey([box('F0EDF6', undefined, 'c=0'), box('F0EDF6', undefined, 'c=2')], 1, 'FFFFFF')
  assert.deepEqual(plan.map(p => [p.appear, p.disappear]), [[null, 1], [1, null]])
})

test('a heavier outline is drawn over the shape only while it applies', () => {
  const heavy = { color: '8E75AD', w: 2, dash: null, alpha: 1 }
  const plan = planKey([box('F0EDF6'), box('E3DAEE', heavy), box('F0EDF6')], 2, 'FFFFFF')
  assert.equal(plan[0].spec.line.w, 1)
  assert.equal(plan[1].spec.name, 'Highlight Outline')
  assert.deepEqual([plan[1].appear, plan[1].disappear], [1, 2])
})

function deck(profile, states) {
  return {
    title: 'Test', author: '', canvas: { w: 960, h: 540 }, profile, footer: 'Footer', themeImages: {}, images: new Map(), equations: new Map(),
    slides: [{ no: 1, layout: 'default', title: 'Test', note: '첫 문단\n\n### 참고\n\n둘째', states, issues: new Map(), snapshots: new Map(), shots: [], mathRects: [] }],
  }
}

for (const profile of [hcs, generic]) {
  test(`the ${profile.name} profile writes a complete package with animations and notes`, () => {
    const title = para('P/div0/h10', 'A title', { role: 'title' })
    const bullet = para('P/div0/ul1/li0', 'A bullet', { li: { type: 'disc', index: 1, listLeft: 26, markerColor: null } })
    const shown = para('P/div0/p2', 'Shown on click')
    const pkg = buildPackage(deck(profile, [{ pageBg: { r: 255, g: 255, b: 255, a: 1 }, items: [title, bullet] }, { pageBg: { r: 255, g: 255, b: 255, a: 1 }, items: [title, bullet, shown] }]))
    const files = new Map(pkg.files)
    for (const part of ['[Content_Types].xml', 'ppt/presentation.xml', 'ppt/slides/slide1.xml', 'ppt/notesSlides/notesSlide1.xml', 'ppt/slideMasters/slideMaster1.xml', 'docProps/core.xml'])
      assert.ok(files.has(part), part)
    const slide = files.get('ppt/slides/slide1.xml')
    assert.match(slide, /<p:ph type="title"\/>/)
    assert.match(slide, /<a:buChar char="•"\/>/)
    assert.match(slide, /presetClass="entr"/)
    assert.match(files.get('ppt/notesSlides/notesSlide1.xml'), /b="1"[^>]*\/><a:t>참고/)
    assert.equal(pkg.info[0].clicks, 1)
  })
}

test('the report lists manual fixes before checks', async t => {
  const dir = await mkdtemp(join(tmpdir(), 'pptx-report-'))
  t.after(() => rm(dir, { recursive: true, force: true }))
  const issues = new Map([['shadow|P/div0', { kind: 'shadow', label: '<div> "Card"', detail: '안쪽 그림자', action: '생략', severity: 'fix', clicks: [0, 1] }]])
  const summary = await writeReport(join(dir, 'deck.report.md'), {
    title: 'Deck', out: 'deck.pptx', slideCount: 1, clickCount: 1, slides: [{ no: 3, issues }], imageFailures: [],
    check: { ran: false, reason: '테스트', flagged: [], threshold: 3 }, notes: ['확인할 것'],
  })
  const md = await readFile(join(dir, 'deck.report.md'), 'utf8')
  assert.equal(summary.fixes, 1)
  assert.ok(md.indexOf('## 직접 고칠 항목') < md.indexOf('## PowerPoint에서 확인할 것'))
  assert.match(md, /### 3쪽 · <div> "Card"/)
  assert.match(md, /클릭 0–1/)
})
