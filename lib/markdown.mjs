// Keep the authored file ordinary Markdown. Presentation styling lives here.
export function citationIds(content = '') {
  const prose = content.replace(/```[\s\S]*?```/g, '').replace(/`[^`]*`/g, '')
  return [...new Set([...prose.matchAll(/(?<!!)\[(\d+(?:,\s*\d+)*)\](?![([])/g)]
    .flatMap(m => m[1].split(',').map(Number)))].sort((a, b) => a - b)
}

export function authoringMarkdown(md, { renderImage } = {}) {
  md.inline.ruler.after('link', 'authoring_citation', (state, silent) => {
    const match = state.src.slice(state.pos).match(/^\[(\d+(?:,\s*\d+)*)\](?![([])/)
    if (!match) return false
    if (!silent) {
      const token = state.push('html_inline', '', 0)
      token.content = `<sup class="authoring-cite">[${match[1].replace(/\s/g, '')}]</sup>`
    }
    state.pos += match[0].length
    return true
  })
  const originalImage = md.renderer.rules.image
  md.renderer.rules.image = (tokens, i, options, env, self) => {
    const token = tokens[i], src = token.attrGet('src') || ''
    if (!src.startsWith('/figures/')) return originalImage(tokens, i, options, env, self)
    const attr = value => md.utils.escapeHtml(value).replaceAll('{', '&#123;')
    const caption = token.content || ''
    const element = renderImage?.({ src, caption, escape:attr })
      ?? `<AuthoringFigure src="${attr(src)}" alt="${attr(caption)}" caption="${attr(caption)}" />`
    return `<div class="authoring-visual">${element}</div>`
  }
  md.core.ruler.after('inline', 'authoring_standalone_images', state => {
    for (let i = 1; i < state.tokens.length - 1; i++) {
      const t = state.tokens[i]
      if (t.type === 'inline' && t.children?.length === 1 && t.children[0].type === 'image'
          && t.children[0].attrGet('src')?.startsWith('/figures/')
          && state.tokens[i - 1].type === 'paragraph_open') {
        state.tokens[i - 1].hidden = true
        state.tokens[i + 1].hidden = true
      }
    }
  })
  // Slidev omits raw slide content from production route metadata. Compile the
  // citation list into the rendered slide so dev, export, and build agree.
  md.core.ruler.after('inline', 'authoring_references', state => {
    const ids = citationIds(state.src)
    if (!ids.length) return
    const token = new state.Token('html_block', '', 0)
    token.content = `<AuthoringReferences :ids='${JSON.stringify(ids)}' />\n`
    state.tokens.push(token)
  })
}
