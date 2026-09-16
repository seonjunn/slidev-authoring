# Slidev Authoring

A small Slidev addon for Markdown citations and editable figure highlights. The theme controls the slide design; the deck owns its figures, notes, bibliography, and annotation data.

Requires Node.js 22.12+ and Slidev 52.19.1. Tested with Slidev 52.19.1, Vue 3.5, and Vite 8.2.2. It does not require the HCS theme.

## Install

```sh
npm install -D github:seonjunn/slidev-authoring
```

Commit `package-lock.json` to pin the package revision. Add this to the deck's first frontmatter block:

```yaml
addons:
  - slidev-addon-authoring
```

Create `vite.config.ts` beside `slides.md`:

```ts
import { fileURLToPath } from 'node:url'
import { authoring } from 'slidev-addon-authoring/vite'

export default authoring({
  root: fileURLToPath(new URL('.', import.meta.url)),
})
```

Put images in `public/figures/`. Create `annotations.json` with `{}` and `references.json` with your bibliography:

```json
{
  "1": {
    "authors": "Author et al.",
    "title": "Paper title",
    "venue": "Conference 2026"
  }
}
```

Use ordinary Markdown:

```md
---
contentId: example
clicks: 1
---

# A concise claim

- Supporting evidence belongs beside its citation. [1]

![Figure caption](/figures/example.png)
```

Numeric citations become superscripts and a bibliography footer. Local `/figures/` images become editable figures. Other images keep Slidev's default rendering. Annotations are stored in source-image percentages and follow the slide's `contentId`, so reordering slides preserves them.

## Edit highlights

Run the deck's Slidev dev server, hover over a figure, and select **그림 강조 편집**. The editor supports multiple regions per click step, drag, resize, keyboard movement, undo/redo, preview, and save. Its current interface is Korean. Saving changes only the deck's annotation JSON. Concurrent edits to the same figure are rejected rather than overwritten.

The editor currently supports one editable figure per `contentId` and up to 40 click states. Set a unique `contentId` and the intended `clicks` in Markdown before editing. Image coordinates must be checked again if the image's crop changes. Mermaid structure editing is outside this addon.

The editor and write endpoint exist only in the local dev server. Static builds include figure rendering and citations without the editor or file-saving endpoint. Slidev's built-in Markdown editor and `/notes-edit` are separate Slidev features.

## Configuration

`authoring()` accepts `root` (required absolute path), `entry`, `annotations`, and `references` (relative paths inside that root). Defaults are `slides.md`, `annotations.json`, and `references.json`.

Optional `renderImage({ src, caption, escape })` returns custom markup for a deck-specific image, or `undefined` to use the standard editable figure. `markdownSetup(md)` adds deck-specific Markdown behavior after the addon. No bibliography or diagram data is bundled into the package.

Use CSS variables to match a theme:

```css
:root {
  --authoring-highlight: #7655ab;
  --authoring-reference-inset: 40px;
  --authoring-reference-bottom: 20px;
  --authoring-reference-size: 11px;
  --authoring-caption-size: 15px;
}
```

`.authoring-visual` fills its parent by default. Set a height for that parent or customize the class in the deck when using a layout without a dedicated figure slot.

## Development

```sh
npm ci
npm test
npm pack
```

The package ships Vue source directly, following Slidev's addon convention. It includes no seminar content, deployment configuration, or mandatory theme dependency.
