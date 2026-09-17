// Optional colors preserve the original four-coordinate annotation format.
export const DEFAULT_HIGHLIGHT = '#7655AB'
export const HIGHLIGHT_PRESETS = [
  { name: '퍼플', color: '#7655AB' },
  { name: '블루', color: '#397DA8' },
  { name: '틸', color: '#21877E' },
  { name: '앰버', color: '#B77A24' },
  { name: '로즈', color: '#B45871' },
]
export const isHighlightColor = value => typeof value === 'string' && /^#[0-9a-f]{6}$/i.test(value)
export function highlightStyle(region, figureColor) {
  const color = isHighlightColor(region?.[4]) ? region[4] : figureColor
  return isHighlightColor(color) ? { '--region-highlight': color } : {}
}
