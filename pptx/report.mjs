// Conversion report: what needs manual work, where the shape check disagrees, and what to look at in PowerPoint.
import { mkdir, writeFile, rm } from 'node:fs/promises'
import { basename, join } from 'node:path'

const clicksText = cs => cs.length === 1 ? `클릭 ${cs[0]}` : `클릭 ${cs[0]}–${cs[cs.length - 1]}`
const pages = nos => [...new Set(nos)].sort((a, b) => a - b).join(', ')

export async function writeReport(path, r) {
  const dir = path.replace(/\.md$/, '')
  const rel = join(basename(dir))
  await rm(dir, { recursive: true, force: true })
  await mkdir(dir, { recursive: true })
  const fixes = [], approx = []
  for (const s of r.slides) {
    const same = new Map()
    for (const is of s.issues.values()) {
      if (is.severity !== 'fix') { approx.push({ no: s.no, ...is }); continue }
      // one entry per kind of problem on a slide; the first few crops show where
      const k = `${is.kind}|${is.detail}`
      if (!same.has(k)) { same.set(k, { no: s.no, ...is, labels: [], crops: [], shapes: [] }); fixes.push(same.get(k)) }
      const g = same.get(k)
      g.labels.push(is.label); if (is.shape) g.shapes.push(is.shape)
      if (is.crop && g.crops.length < 3) g.crops.push(is.crop)
      g.clicks = [...new Set([...g.clicks, ...is.clicks])].sort((a, b) => a - b)
    }
  }
  for (const f of r.imageFailures) fixes.push({ no: null, label: f, detail: '그림 파일을 가져오지 못했습니다.', action: '생략', clicks: [0] })

  // the same approximation is one line, listing its slides
  const groups = new Map()
  for (const a of approx) {
    const k = `${a.kind}|${a.detail}`
    if (!groups.has(k)) groups.set(k, { ...a, labels: [], nos: [] })
    groups.get(k).labels.push(a.label)
    groups.get(k).nos.push(a.no)
  }

  const L = []
  L.push(`# PowerPoint 변환 보고서`, '')
  L.push(`**${r.title || r.out}** → \`${r.out}\` · 슬라이드 ${r.slideCount}장 · 클릭 ${r.clickCount}회`, '')
  L.push('## 요약', '')
  L.push(`- **직접 고칠 항목: ${fixes.reduce((a, f) => a + (f.labels?.length || 1), 0)}개**${fixes.length ? ' — 아래 목록대로 PowerPoint에서 고쳐야 합니다.' : ''}`)
  if (r.check.ran) L.push(`- **모양 비교에서 의심되는 곳: ${r.check.flagged.length}곳** — LibreOffice로 그린 결과가 원본과 ${r.check.threshold}% 넘게 다른 클릭 단계입니다.`)
  else L.push(`- **모양 비교를 하지 않았습니다.** ${r.check.reason} 글자 넘침·줄바꿈 차이는 확인되지 않았으니 PowerPoint에서 모든 슬라이드를 훑어보세요.`)
  L.push(`- **PowerPoint에서 확인할 것: ${r.notes.length + groups.size}가지** — 근사로 옮긴 부분과 LibreOffice로는 확인할 수 없는 부분입니다.`, '')

  L.push('## 직접 고칠 항목', '')
  if (!fixes.length) L.push('없습니다.', '')
  let n = 0
  for (const f of fixes) {
    const labels = f.labels || [f.label]
    L.push(`### ${f.no ? `${f.no}쪽` : '그림 파일'} · ${labels.length > 1 ? `${labels[0]} 외 ${labels.length - 1}개` : labels[0]}`, '')
    L.push(`- 문제: ${f.detail}`, `- 처리: ${f.action}${f.shapes?.length ? ` (PowerPoint 도형 이름: ${f.shapes.map(s => `\`${s}\``).join(', ')})` : ''}`)
    if (labels.length > 1) L.push(`- 해당 요소: ${[...new Set(labels)].slice(0, 8).join(', ')}${new Set(labels).size > 8 ? ' …' : ''}`)
    if (f.no) L.push(`- 나타나는 단계: ${clicksText(f.clicks)}`)
    for (const crop of f.crops || []) { n++; const file = `fix-${String(n).padStart(3, '0')}.png`; await writeFile(join(dir, file), crop); L.push('', `![원본 화면](${rel}/${file})`) }
    L.push('')
  }

  L.push('## 모양 비교에서 의심되는 곳', '')
  if (!r.check.ran) L.push(`비교하지 않았습니다. ${r.check.reason}`, '')
  else if (!r.check.flagged.length) L.push('없습니다. LibreOffice로 그린 모든 클릭 단계가 원본과 가깝습니다. PowerPoint와 LibreOffice는 글꼴을 다르게 그릴 수 있으니 아래 확인 목록도 보세요.', '')
  else {
    L.push('왼쪽이 원본, 오른쪽이 변환 결과(LibreOffice)입니다. LibreOffice와 PowerPoint의 차이 때문에 생긴 표시일 수도 있습니다.', '')
    for (const [i, f] of r.check.flagged.entries()) {
      const file = `diff-${String(i + 1).padStart(3, '0')}.png`
      await writeFile(join(dir, file), f.crop)
      L.push(`### ${f.no}쪽 · 클릭 ${f.click} · 차이 ${f.score.toFixed(1)}%`, '', `![원본 | 변환 결과](${rel}/${file})`, '')
    }
  }

  L.push('## PowerPoint에서 확인할 것', '')
  for (const note of r.notes) L.push(`- ${note}`)
  for (const g of groups.values()) L.push(`- ${g.detail} ${g.labels.length > 1 ? `${g.labels[0]} 외 ${g.labels.length - 1}개` : g.labels[0]} (${pages(g.nos)}쪽, ${g.action})`)
  if (!r.notes.length && !groups.size) L.push('없습니다.')
  L.push('')

  L.push('## agent에게 넘길 때', '')
  L.push('이 보고서와 PPTX 파일을 함께 주고 "직접 고칠 항목과 의심되는 곳을 보고서의 원본 이미지에 맞게 PPTX에서 고쳐줘"라고 요청하세요. 도형 이름은 PowerPoint 선택 창에 보이는 이름과 같습니다.', '')
  await writeFile(path, L.join('\n'))
  return { fixes: fixes.reduce((a, f) => a + (f.labels?.length || 1), 0), flagged: r.check.ran ? r.check.flagged.length : null, notes: r.notes.length + groups.size }
}

// Things the converter knows it approximated, and what LibreOffice cannot show.
export function collectNotes({ slides, info, fonts, rasterized }) {
  const notes = []
  const clicked = slides.filter(s => s.states.length > 1)
  if (clicked.length) notes.push(`클릭 애니메이션은 실제로 재생해 확인하지 않았습니다(${pages(clicked.map(s => s.no))}쪽). 모양 비교는 클릭마다 도착할 모습을 계산해 비교한 것이니, 발표 전에 PowerPoint 슬라이드 쇼로 클릭을 끝까지 넘겨 보세요.`)
  const byNo = new Map(info.map(i => [i.no, i]))
  const copies = slides.filter(s => byNo.get(s.no)?.copies.length)
  if (copies.length) notes.push(`클릭마다 글 내용이 바뀌는 도형은 같은 자리에 겹친 복사본입니다(${copies.reduce((a, s) => a + byNo.get(s.no).copies.length, 0)}개, ${pages(copies.map(s => s.no))}쪽). 글을 고칠 때 겹친 복사본도 함께 고치세요.`)
  const lost = slides.filter(s => byNo.get(s.no)?.lostClicks.length)
  if (lost.length) notes.push(`화면 변화가 없는 클릭을 PowerPoint 애니메이션으로 남기지 못해 클릭 수가 원본보다 적습니다(${pages(lost.map(s => s.no))}쪽). 대본의 클릭 표시와 맞춰 보세요.`)
  const math = slides.filter(s => s.states.some(st => st.items.some(it => it.kind === 'para' && it.runs.some(r => r.math !== undefined))))
  if (math.length) notes.push(`수식은 PowerPoint 수식으로 넣었습니다(${pages(math.map(s => s.no))}쪽). LibreOffice는 이 수식을 그리지 못해 모양 비교에서 제외했으니 PowerPoint에서 확인하세요.`)
  const hangul = slides.filter(s => s.states.at(-1).items.some(it => it.kind === 'para' && it.runs.some(r => /[가-힣]/.test(r.t || ''))))
  if (hangul.length) notes.push(`한글 본문이 있습니다(${pages(hangul.map(s => s.no))}쪽). PowerPoint의 한글 글꼴은 폭이 달라 줄바꿈이 바뀔 수 있으니 글자 넘침을 확인하세요.`)
  const subs = Object.entries(fonts).sort((a, b) => b[1] - a[1])
  if (subs.length) notes.push(`글꼴을 PowerPoint에 있는 글꼴로 바꿨습니다: ${subs.map(([k, v]) => `${k} (${v}곳)`).join(', ')}.`)
  if (rasterized.length) notes.push(`PNG·JPEG·GIF가 아닌 그림 ${rasterized.length}개는 PNG로 바꿔 넣었습니다.`)
  return notes
}
