import { resolve, isAbsolute, relative } from 'node:path'
import { existsSync } from 'node:fs'
import { authoringMarkdown } from './lib/markdown.mjs'
import { annotationEditor } from './lib/server.mjs'

// The deck owns its data and presentation-specific Markdown extensions.
export function authoring({ root, entry = 'slides.md', annotations = 'annotations.json', references = 'references.json', renderImage, markdownSetup } = {}) {
  if (!root || !isAbsolute(root)) throw new Error('authoring.root must be an absolute deck directory')
  for (const file of [entry, annotations, references]) {
    const rel = relative(root, resolve(root, file))
    if (isAbsolute(rel) || rel === '..' || rel.startsWith('../')) throw new Error('Authoring files must stay inside the deck directory')
  }
  const id = 'virtual:slidev-authoring/data', resolved = '\0' + id
  const dataFiles = [annotations, references].map(file => resolve(root, file))
  return {
    slidev: { markdown: { markdownSetup(md) {
      authoringMarkdown(md, { renderImage })
      markdownSetup?.(md)
    } } },
    plugins: [{
      name:'slidev-authoring-data',
      resolveId(source) { if (source === id) return resolved },
      load(source) {
        if (source !== resolved) return
        return dataFiles.map((file, i) => {
          this.addWatchFile(file)
          const name = i ? 'bibliography' : 'annotations'
          return existsSync(file) ? `export { default as ${name} } from ${JSON.stringify(file)}` : `export const ${name} = {}`
        }).join('\n')
      },
      handleHotUpdate({ file, server }) {
        if (dataFiles.includes(file)) {
          const mod = server.moduleGraph.getModuleById(resolved)
          if (mod) server.moduleGraph.invalidateModule(mod)
        }
      },
    }, annotationEditor(root, { entry, annotations })],
  }
}
