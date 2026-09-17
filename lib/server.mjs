import { readFile, writeFile, rename } from 'node:fs/promises'
import { resolve } from 'node:path'
import { createHash } from 'node:crypto'
import { parse } from '@slidev/parser'
import { isHighlightColor } from './colors.mjs'

const revisionOf = entry => createHash('sha256').update(JSON.stringify(entry ?? null)).digest('hex')
const fail = (status, message) => { throw Object.assign(new Error(message), { status }) }

export function validateRegions(regions, steps) {
  if (!Array.isArray(regions) || regions.length !== steps) fail(400, 'The slide click count changed. Reopen the editor.')
  for (const stage of regions) {
    if (!Array.isArray(stage) || stage.length > 40) fail(400, 'Too many highlight regions.')
    for (const r of stage) {
      if (!Array.isArray(r) || ![4, 5].includes(r.length) || r.slice(0, 4).some(n => typeof n !== 'number' || !Number.isFinite(n))) fail(400, 'Invalid region coordinates.')
      if (r.length === 5 && !isHighlightColor(r[4])) fail(400, 'Use a six-digit hex highlight color, such as #7655AB.')
      const [x, y, w, h] = r
      if (x < -20 || y < -20 || w <= 0 || h <= 0 || x + w > 120 || y + h > 120) fail(400, 'A region is outside the figure bounds.')
    }
  }
}

// Only this deck's fixed annotation file can be changed. A revision check
// prevents one open editor from silently overwriting another editor's work.
export function createAnnotationStore(root, { entry = "slides.md", annotations = "annotations.json" } = {}) {
  const file = resolve(root, annotations)
  let queue = Promise.resolve()
  const load = async () => JSON.parse(await readFile(file, 'utf8').catch(error => { if (error.code === 'ENOENT') return '{}'; throw error }))
  async function slide(slideId, src) {
    const data = await parse(await readFile(resolve(root, entry), 'utf8'))
    const matches = data.slides.filter(s => s.frontmatter.contentId === slideId)
    if (matches.length !== 1) fail(404, 'Slide not found. Reopen the editor.')
    const item = matches[0]
    const sources = [...item.content.matchAll(/!\[[^\n]*?\]\((\/figures\/[^\s)]+)\)/g)].map(m => m[1])
    if (!sources.includes(src)) fail(409, 'The figure changed in Markdown. Reopen the editor.')
    const steps = Number(item.frontmatter.clicks || 0) + 1
    if (!Number.isInteger(steps) || steps < 1 || steps > 40) fail(400, 'Unsupported click count.')
    return { steps }
  }
  return {
    async get(slideId, src) {
      const { steps } = await slide(slideId, src)
      const all = await load(), stored = all[slideId]
      return { revision: revisionOf(stored), entry: stored?.src === src ? stored : null, steps }
    },
    save(payload) {
      const operation = queue.then(async () => {
        const { slideId, src, revision, regions, focusStyle, focusColor, ratio } = payload
        const { steps } = await slide(slideId, src)
        validateRegions(regions, steps)
        if (focusColor !== undefined && !isHighlightColor(focusColor)) fail(400, 'Use a six-digit hex highlight color, such as #7655AB.')
        if (!['wash', 'frame'].includes(focusStyle) || !Number.isFinite(ratio) || ratio <= 0 || ratio > 100) fail(400, 'Invalid figure style.')
        const all = await load()
        if (revision !== revisionOf(all[slideId])) fail(409, 'Another editor saved this figure. Reopen it before saving again.')
        all[slideId] = { src, ratio, focusStyle, ...(focusColor ? { focusColor } : {}), regions }
        const temporary = file + '.tmp'
        await writeFile(temporary, JSON.stringify(all, null, 2) + '\n')
        await rename(temporary, file)
        return { revision: revisionOf(all[slideId]), entry: all[slideId], steps }
      })
      queue = operation.catch(() => {})
      return operation
    },
  }
}

export function annotationEditor(root, options = {}) {
  const store = createAnnotationStore(root, options)
  return {
    name: 'slidev-authoring-editor',
    apply: 'serve',
    configureServer(server) {
      server.middlewares.use('/__slidev-authoring/annotations', async (req, res) => {
        res.setHeader('Content-Type', 'application/json; charset=utf-8')
        res.setHeader('Cache-Control', 'no-store')
        try {
          if (req.method === 'GET') {
            if (req.headers['sec-fetch-site'] === 'cross-site') fail(403, 'Open this editor from the local slide server.')
            const params = new URL(req.url, 'http://localhost').searchParams
            res.end(JSON.stringify(await store.get(params.get('slideId'), params.get('src'))))
          } else if (req.method === 'POST') {
            const origin = req.headers.origin
            if (!origin || new URL(origin).host !== req.headers.host || !req.headers['content-type']?.startsWith('application/json')) fail(403, 'Save from the local slide editor.')
            let body = ''
            for await (const chunk of req) {
              body += chunk
              if (Buffer.byteLength(body) > 100_000) fail(413, 'The annotation request is too large.')
            }
            res.end(JSON.stringify(await store.save(JSON.parse(body))))
          } else fail(405, 'Method not allowed.')
        } catch (error) {
          res.statusCode = error.status || 500
          res.end(JSON.stringify({ error: error.status ? error.message : 'Could not save the annotations. Check the local server.' }))
        }
      })
    },
  }
}
