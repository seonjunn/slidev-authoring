// Shape check: draw the converted slides with LibreOffice and compare each click state with the browser.
import { execFile } from 'node:child_process'
import { promisify } from 'node:util'
import { existsSync } from 'node:fs'
import { readdir, readFile } from 'node:fs/promises'
import { join, dirname, basename } from 'node:path'
import { fileURLToPath } from 'node:url'
import { compareImages, sideBySide } from './page.mjs'

const run = promisify(execFile)
const IMAGE = 'slidev-authoring-libreoffice:1'
const HERE = dirname(fileURLToPath(import.meta.url))
const ok = async (cmd, args) => { try { await run(cmd, args, { timeout: 20000 }); return true } catch { return false } }

async function which(names) {
  for (const n of names) {
    if (n.startsWith('/')) { if (existsSync(n)) return n; continue }
    try { const { stdout } = await run('sh', ['-c', `command -v ${n}`]); if (stdout.trim()) return stdout.trim() } catch {}
  }
  return null
}

// LibreOffice on this machine, else in Docker; null with the reason when neither is usable.
export async function findRenderer() {
  const soffice = await which(['soffice', 'libreoffice', '/Applications/LibreOffice.app/Contents/MacOS/soffice'])
  const pdftoppm = await which(['pdftoppm'])
  const docker = await which(['docker']) && await ok('docker', ['info', '--format', '{{.ServerVersion}}'])
  if (soffice && pdftoppm) return { kind: 'local', soffice, pdftoppm, docker }
  if (docker) return { kind: 'docker' }
  return { kind: null, reason: soffice ? 'pdftoppm(poppler)이 없어' : 'LibreOffice와 Docker를 찾지 못해' }
}

async function ensureImage(log) {
  if (await ok('docker', ['image', 'inspect', IMAGE])) return
  log('LibreOffice Docker 이미지를 처음 만듭니다. 몇 분 걸립니다.')
  await run('docker', ['build', '-t', IMAGE, '-f', join(HERE, 'libreoffice.Dockerfile'), HERE], { timeout: 3600000, maxBuffer: 64 << 20 })
}

// PPTX → one PNG per slide (1 CSS px per pixel), in slide order.
export async function renderSlides(pptx, dir, renderer, log) {
  const name = basename(pptx, '.pptx')
  const local = async () => {
    await run(renderer.soffice, [`-env:UserInstallation=file://${dir}/lo-profile`, '--headless', '--convert-to', 'pdf', '--outdir', dir, pptx], { timeout: 900000 })
    if (!existsSync(join(dir, name + '.pdf'))) throw new Error('LibreOffice did not write a PDF')
    await run(renderer.pdftoppm, ['-png', '-r', '96', join(dir, name + '.pdf'), join(dir, 'page')], { timeout: 900000 })
  }
  const docker = async () => {
    await ensureImage(log)
    const uid = process.getuid?.() ?? 1000, gid = process.getgid?.() ?? 1000
    await run('docker', ['run', '--rm', '-u', `${uid}:${gid}`, '-e', 'HOME=/tmp', '-v', `${dirname(pptx)}:/in:ro`, '-v', `${dir}:/out`, IMAGE, 'sh', '-c',
      `soffice -env:UserInstallation=file:///tmp/lo --headless --convert-to pdf --outdir /out "/in/${name}.pptx" && pdftoppm -png -r 96 "/out/${name}.pdf" /out/page`], { timeout: 1800000, maxBuffer: 64 << 20 })
  }
  if (renderer.kind === 'local') {
    try { await local() } catch (e) { if (!renderer.docker) throw e; log('로컬 LibreOffice로 변환하지 못해 Docker로 다시 시도합니다.'); await docker() }
  } else await docker()
  const pages = (await readdir(dir)).filter(f => /^page-\d+\.png$/.test(f)).sort((a, b) => parseInt(a.slice(5)) - parseInt(b.slice(5)))
  return pages.map(f => join(dir, f))
}

const dataUrl = buf => 'data:image/png;base64,' + buf.toString('base64')

// states: [{ no, click, shot (Buffer), masks }]; renders: PNG paths in the same order
export async function compareStates(page, states, renders, canvas, threshold) {
  const flagged = []
  for (let i = 0; i < states.length; i++) {
    const s = states[i]
    if (!s.shot || !renders[i]) continue
    const b = await readFile(renders[i])
    const { score, box } = await page.evaluate(compareImages, { a: dataUrl(s.shot), b: dataUrl(b), masks: s.masks, width: canvas.w, height: canvas.h })
    if (score > threshold && box) {
      const crop = await page.evaluate(sideBySide, { a: dataUrl(s.shot), b: dataUrl(b), box, width: canvas.w, height: canvas.h })
      flagged.push({ no: s.no, click: s.click, score, box, crop: Buffer.from(crop.split(',')[1], 'base64') })
    }
  }
  return flagged
}
