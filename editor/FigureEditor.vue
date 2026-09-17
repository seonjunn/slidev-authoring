<script setup>
import { computed, nextTick, onBeforeUnmount, onMounted, ref } from 'vue'
import { DEFAULT_HIGHLIGHT, HIGHLIGHT_PRESETS, isHighlightColor, highlightStyle } from '../lib/colors.mjs'
const props = defineProps({ slideId:String, src:String, href:String, caption:String, ratio:Number, initialStep:Number })
const emit = defineEmits(['close', 'saved'])
const dialog = ref(), svg = ref(), loading = ref(true), saving = ref(false), error = ref('')
const step = ref(0), selected = ref(-1), mode = ref('select'), style = ref('wash'), regions = ref([[]])
const focusColor = ref(''), themeColor = ref(DEFAULT_HIGHLIGHT)
const currentColor = computed(() => boxes.value[selected.value]?.[4] || focusColor.value || themeColor.value)
const inheritsColor = computed(() => selected.value >= 0 ? !boxes.value[selected.value]?.[4] : !focusColor.value)
const revision = ref(''), baseline = ref(''), history = ref([]), future = ref([]), preview = ref(false)
let gesture = null, restoreFocus, previousOverflow
const state = () => JSON.stringify({ regions:regions.value, focusStyle:style.value, focusColor:focusColor.value })
const dirty = computed(() => !loading.value && state() !== baseline.value)
const boxes = computed(() => regions.value[step.value] || [])
const endpoint = '/__slidev-authoring/annotations', minimum = .25
const handles = ['nw', 'ne', 'sw', 'se']
const clone = value => JSON.parse(JSON.stringify(value))
function remember(snapshot = state()) { history.value.push(snapshot); future.value = [] }
function restore(snapshot) {
  const value = JSON.parse(snapshot)
  regions.value = value.regions; style.value = value.focusStyle; focusColor.value = value.focusColor || ''; selected.value = -1
}
function undo() { if (history.value.length) { future.value.push(state()); restore(history.value.pop()) } }
function redo() { if (future.value.length) { history.value.push(state()); restore(future.value.pop()) } }
function changeStyle(value) { remember(); style.value = value }
function changeColor(value) {
  if (!isHighlightColor(value)) { error.value = '색상은 #RRGGBB 형식으로 입력하세요.'; return }
  error.value = ''
  const color = value.toUpperCase()
  if (selected.value >= 0 ? boxes.value[selected.value][4] === color : focusColor.value === color) return
  remember()
  if (selected.value >= 0) boxes.value[selected.value][4] = color
  else focusColor.value = color
}
function resetColor() {
  if (inheritsColor.value) return
  remember(); error.value = ''
  if (selected.value >= 0) boxes.value[selected.value].splice(4)
  else focusColor.value = ''
}
function setRect(index, coordinates) { boxes.value[index] = [...coordinates, ...boxes.value[index].slice(4)] }
function remove() {
  if (selected.value < 0) return
  remember(); regions.value[step.value].splice(selected.value, 1); selected.value = -1
}
function duplicate() {
  if (selected.value < 0 || boxes.value.length >= 40) return
  remember(); boxes.value.push(clone(boxes.value[selected.value])); selected.value = boxes.value.length - 1
}
function changeStep(value) { step.value = value; selected.value = -1; mode.value = 'select' }
function position(event) {
  const p = new DOMPoint(event.clientX, event.clientY).matrixTransform(svg.value.getScreenCTM().inverse())
  return [Math.max(-1.5, Math.min(101.5, p.x)), Math.max(-1.5 * props.ratio, Math.min(100 + 1.5 * props.ratio, p.y * props.ratio))]
}
function start(event, index = -1, handle = null) {
  if (preview.value || loading.value || saving.value || event.button !== 0) return
  event.preventDefault(); event.stopPropagation()
  const at = position(event), snapshot = state()
  if (mode.value === 'draw') {
    if (boxes.value.length >= 40) return
    boxes.value.push([at[0], at[1], minimum, minimum]); index = boxes.value.length - 1
    gesture = { type:'draw', at, snapshot, index }
  } else if (index >= 0) gesture = { type:handle || 'move', at, snapshot, index, rect:[...boxes.value[index]] }
  else { selected.value = -1; return }
  selected.value = index; svg.value.setPointerCapture(event.pointerId)
}
function move(event) {
  if (!gesture) return
  const [x, y] = position(event), g = gesture, [sx, sy] = g.at
  if (g.type === 'draw') setRect(g.index, [Math.min(x, sx), Math.min(y, sy), Math.max(minimum, Math.abs(x - sx)), Math.max(minimum, Math.abs(y - sy))])
  else {
    const [a, b, w, h] = g.rect
    if (g.type === 'move') setRect(g.index, [Math.max(-1.5, Math.min(101.5 - w, a + x - sx)), Math.max(-1.5 * props.ratio, Math.min(100 + 1.5 * props.ratio - h, b + y - sy)), w, h])
    else {
      const left = g.type.includes('w') ? Math.min(a + w - minimum, x) : a
      const right = g.type.includes('e') ? Math.max(a + minimum, x) : a + w
      const top = g.type.includes('n') ? Math.min(b + h - minimum, y) : b
      const bottom = g.type.includes('s') ? Math.max(b + minimum, y) : b + h
      setRect(g.index, [left, top, right - left, bottom - top])
    }
  }
}
function end(event) {
  if (!gesture) return
  const snapshot = gesture.snapshot
  if (event.type === 'pointercancel') restore(snapshot)
  else if (snapshot !== state()) remember(snapshot)
  gesture = null; mode.value = 'select'
  if (svg.value.hasPointerCapture(event.pointerId)) svg.value.releasePointerCapture(event.pointerId)
}
function handlePoint(r, h) { return [r[0] + (h.includes('e') ? r[2] : 0), (r[1] + (h.includes('s') ? r[3] : 0)) / props.ratio] }
function keydown(event) {
  event.stopPropagation()
  if (event.key === 'Tab') {
    const controls = [...dialog.value.querySelectorAll('button:not(:disabled),select:not(:disabled),input:not(:disabled)')]
    const first = controls[0], last = controls.at(-1)
    if (event.shiftKey && document.activeElement === first) { event.preventDefault(); last?.focus() }
    else if (!event.shiftKey && document.activeElement === last) { event.preventDefault(); first?.focus() }
    return
  }
  if (event.key === 'Escape') { event.preventDefault(); if (dirty.value) error.value = '저장하거나 “변경 버리기”를 눌러 닫으세요.'; else emit('close'); return }
  if (/^(INPUT|SELECT|TEXTAREA)$/.test(event.target.tagName)) return
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 'z') { event.preventDefault(); event.shiftKey ? redo() : undo(); return }
  if ((event.metaKey || event.ctrlKey) && event.key.toLowerCase() === 's') { event.preventDefault(); save(); return }
  if (preview.value || selected.value < 0 || saving.value) return
  if (['Delete', 'Backspace'].includes(event.key)) { event.preventDefault(); remove(); return }
  const delta = {ArrowLeft:[-1,0], ArrowRight:[1,0], ArrowUp:[0,-1], ArrowDown:[0,1]}[event.key]
  if (delta) {
    event.preventDefault(); remember()
    const r = boxes.value[selected.value], amount = event.shiftKey ? 1 : .2
    r[0] = Math.max(-1.5, Math.min(101.5-r[2], r[0]+delta[0]*amount))
    r[1] = Math.max(-1.5*props.ratio, Math.min(100+1.5*props.ratio-r[3], r[1]+delta[1]*amount))
  }
}
async function save() {
  if (loading.value || saving.value || !dirty.value) return
  saving.value = true; error.value = ''
  try {
    const payload = { slideId:props.slideId, src:props.src, revision:revision.value, ratio:props.ratio, focusStyle:style.value, ...(focusColor.value ? {focusColor:focusColor.value} : {}),
      regions:regions.value.map(stage => stage.map(r => r.map((n, i) => i < 4 ? Math.round(n*1000)/1000 : n))) }
    const response = await fetch(endpoint, { method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify(payload) })
    const result = await response.json()
    if (!response.ok) throw new Error(result.error)
    baseline.value = state(); emit('saved', result.entry)
  } catch (e) { error.value = e.message }
  finally { saving.value = false }
}
function unload(event) { if (dirty.value) { event.preventDefault(); event.returnValue = '' } }
onMounted(async () => {
  const themed = getComputedStyle(document.documentElement).getPropertyValue('--authoring-highlight').trim()
  if (isHighlightColor(themed)) themeColor.value = themed
  restoreFocus = document.activeElement; previousOverflow = document.body.style.overflow
  document.body.style.overflow = 'hidden'; window.addEventListener('beforeunload', unload)
  await nextTick(); dialog.value.focus()
  try {
    const response = await fetch(`${endpoint}?${new URLSearchParams({slideId:props.slideId, src:props.src})}`)
    const result = await response.json()
    if (!response.ok) throw new Error(result.error)
    revision.value = result.revision
    const original = result.entry?.regions || [[]]
    regions.value = Array.from({length:result.steps}, (_, i) => clone(original[Math.min(i, original.length-1)] || []))
    focusColor.value = result.entry?.focusColor || ''
    style.value = result.entry?.focusStyle || 'wash'; step.value = Math.min(props.initialStep || 0, result.steps-1)
    baseline.value = state(); loading.value = false
  } catch (e) { error.value = e.message }
})
onBeforeUnmount(() => { window.removeEventListener('beforeunload', unload); document.body.style.overflow = previousOverflow; restoreFocus?.focus() })
</script>

<template>
  <div class="figure-editor-backdrop" @pointerdown.stop @click.stop>
    <section ref="dialog" class="figure-editor" role="dialog" aria-modal="true" aria-labelledby="figure-editor-title" tabindex="-1" @keydown="keydown">
      <header><div><h2 id="figure-editor-title">그림 강조 편집</h2><p>{{ caption }}</p></div><span>{{ dirty ? '저장하지 않은 변경' : '원본 파일에 저장' }}</span></header>
      <div class="editor-toolbar">
        <label>클릭 단계 <select :value="step" :disabled="loading || saving" @change="changeStep(Number($event.target.value))"><option v-for="(_, i) in regions" :key="i" :value="i">{{ i === 0 ? '첫 화면' : `클릭 ${i}` }}</option></select></label>
        <button :class="{chosen:mode==='select' && !preview}" :disabled="loading || saving" @click="mode='select'; preview=false">선택·이동</button>
        <button :class="{chosen:mode==='draw' && !preview}" :disabled="loading || saving" @click="mode='draw'; preview=false; selected=-1">영역 추가</button>
        <button :disabled="!history.length || saving" @click="undo">실행 취소</button>
        <button :disabled="!future.length || saving" @click="redo">다시 실행</button>
        <label class="preview-label"><input v-model="preview" type="checkbox" /> 발표 화면 미리보기</label>
      </div>
      <div class="editor-workspace">
        <div class="editor-canvas" :class="{drawing:mode==='draw' && !preview}">
          <svg ref="svg" :viewBox="`-1.5 -1.5 103 ${100 / ratio + 3}`" role="img" :aria-label="caption" @pointerdown="start($event)" @pointermove="move" @pointerup="end" @pointercancel="end">
            <image :href="href" x="0" y="0" width="100" :height="100 / ratio" preserveAspectRatio="none" />
            <g v-for="(r, i) in boxes" :key="i">
              <rect class="editor-highlight" :class="{frame:style==='frame'}" :style="highlightStyle(r, focusColor)" :x="r[0]" :y="r[1]/ratio" :width="r[2]" :height="r[3]/ratio" @pointerdown="start($event, i)" />
              <rect v-if="!preview" class="editor-selection" :class="{selected:i===selected}" :x="r[0]" :y="r[1]/ratio" :width="r[2]" :height="r[3]/ratio" @pointerdown="start($event, i)" />
              <circle v-for="h in (!preview && i===selected ? handles : [])" :key="h" class="editor-handle" :cx="handlePoint(r,h)[0]" :cy="handlePoint(r,h)[1]" r=".55" @pointerdown="start($event,i,h)" />
            </g>
          </svg>
          <span v-if="loading" class="editor-loading">{{ error ? '그림을 불러오지 못했습니다.' : '불러오는 중…' }}</span>
        </div>
        <aside>
          <label>강조 모양 <select :value="style" :disabled="loading || saving" @change="changeStyle($event.target.value)"><option value="wash">반투명 배경</option><option value="frame">테두리</option></select></label>
          <fieldset class="color-controls" :disabled="loading || saving">
            <legend>{{ selected >= 0 ? `영역 ${selected + 1} 색상` : '그림 기본색' }}</legend>
            <div class="color-presets">
              <button v-for="preset in HIGHLIGHT_PRESETS" :key="preset.color" :aria-label="`${preset.name} 강조색`" :aria-pressed="!inheritsColor && currentColor.toUpperCase() === preset.color" :class="{chosen:!inheritsColor && currentColor.toUpperCase() === preset.color}" @click="changeColor(preset.color)"><i :style="{background:preset.color}" />{{ preset.name }}</button>
            </div>
            <label class="custom-color">사용자 지정</label>
            <div class="color-inputs">
              <input type="color" aria-label="사용자 지정 색상 선택" :value="currentColor" @change="changeColor($event.target.value)" />
              <input type="text" aria-label="강조색 HEX" :value="currentColor" maxlength="7" spellcheck="false" @change="changeColor($event.target.value)" />
            </div>
            <button class="inherit-color" :class="{chosen:inheritsColor}" @click="resetColor">{{ selected >= 0 ? '그림 기본색 사용' : '테마 색상 사용' }}</button>
            <p>{{ selected >= 0 ? '이 클릭 단계의 선택 영역에만 적용합니다.' : '모든 단계에서 별도 색상이 없는 영역에 적용합니다.' }}</p>
            <button v-if="selected >= 0" @click="selected=-1">그림 기본색 편집</button>
          </fieldset>
          <h3>이 단계의 강조 {{ boxes.length }}개</h3>
          <p v-if="!boxes.length">강조가 없는 단계입니다.<br/>“영역 추가”를 누르고 그림 위를 드래그하세요.</p>
          <div class="region-list"><button v-for="(r, i) in boxes" :key="i" :class="{chosen:selected===i}" @click="selected=i; mode='select'; preview=false"><i class="region-swatch" :style="{background:r[4] || focusColor || themeColor}" />영역 {{ i+1 }}</button></div>
          <div class="selection-actions"><button :disabled="selected<0 || saving" @click="duplicate">복제</button><button :disabled="selected<0 || saving" @click="remove">삭제</button></div>
          <p>영역을 드래그하면 이동합니다. 모서리를 잡으면 크기가 바뀝니다.</p>
          <p>방향키: 미세 이동<br/>Shift + 방향키: 큰 이동<br/>Delete: 선택 영역 삭제</p>
          <p>클릭 수는 슬라이드의 <code>clicks</code> 설정을 따릅니다. 단계마다 여러 영역을 지정할 수 있습니다.</p>
        </aside>
      </div>
      <footer><p role="status" :class="{error}">{{ error || '저장하면 강조가 슬라이드에 바로 반영됩니다.' }}</p><button :disabled="saving" @click="emit('close')">{{ dirty ? '변경 버리기' : '닫기' }}</button><button class="save" :disabled="!dirty || saving || loading" @click="save">{{ saving ? '저장 중…' : '저장' }}</button></footer>
    </section>
  </div>
</template>

<style scoped>
.figure-editor-backdrop { position:fixed;inset:0;z-index:10000;background:#17192480;display:grid;place-items:center;color:#33394b;font:14px/1.5 Arial,'Noto Sans KR',sans-serif; }
.figure-editor { width:min(1240px,96vw);height:min(850px,94vh);background:#fff;border:1px solid #c6c7d1;border-radius:10px;display:flex;flex-direction:column;box-shadow:0 12px 50px #1115;outline:none; }
header { display:flex;gap:24px;justify-content:space-between;align-items:center;padding:16px 20px;border-bottom:1px solid #e0e1e8; }
header h2 { margin:0;font-size:20px;font-weight:600;color:#424769; }
header p { margin:4px 0 0;color:#70778b;font-size:12px;max-width:820px; }
header>span { white-space:nowrap;font-size:12px;color:#70778b; }
.editor-toolbar { display:flex;gap:8px;align-items:center;padding:12px 20px;flex-wrap:wrap;border-bottom:1px solid #e0e1e8; }
label { display:flex;align-items:center;gap:8px; }
button,select { border:1px solid #c8cbd6;border-radius:5px;padding:6px 10px;background:#fff;color:#424769;font:inherit;cursor:pointer; }
button:hover:not(:disabled) { background:#f0edf7; }
button:disabled,select:disabled { opacity:.4;cursor:default; }
button.chosen { background:#eee8f6;border-color:#8067ae;color:#513878; }
button:focus-visible,select:focus-visible,input:focus-visible { outline:2px solid #8067ae;outline-offset:2px; }
.preview-label { margin-left:auto;font-size:13px; }
.editor-workspace { display:grid;grid-template-columns:minmax(0,1fr) 220px;flex:1;min-height:0; }
.editor-canvas { margin:18px;position:relative;min-height:0;min-width:0;background:#fff;border:1px solid #dfe1e9; }
.editor-canvas svg { width:100%;height:100%;display:block;touch-action:none;user-select:none; }
.editor-canvas.drawing { cursor:crosshair; }
.editor-highlight { fill:var(--region-highlight,var(--authoring-highlight,#7655ab));fill-opacity:.19;stroke:none;mix-blend-mode:multiply;rx:.45; }
.editor-highlight.frame { fill:none;stroke:var(--region-highlight,var(--authoring-highlight,#7655ab));stroke-width:2.2px;vector-effect:non-scaling-stroke; }
.editor-selection { fill:transparent;stroke:#9181ac;stroke-width:1px;stroke-dasharray:4 3;vector-effect:non-scaling-stroke;cursor:move; }
.drawing .editor-selection { cursor:crosshair; }
.editor-selection.selected { stroke:#4d356f;stroke-width:1.5px;stroke-dasharray:none; }
.editor-handle { fill:#fff;stroke:#4d356f;stroke-width:1.5px;vector-effect:non-scaling-stroke;cursor:nwse-resize; }
.editor-loading { position:absolute;inset:0;display:grid;place-items:center;background:#ffffffe0; }
aside { padding:18px 16px;border-left:1px solid #e0e1e8;overflow:auto;background:#fafafe; }
aside label { display:block; }aside select { display:block;margin-top:7px;width:100%; }
.color-controls { border:0;padding:0;margin:20px 0 0;min-width:0; }
.color-controls legend { font-weight:600;margin-bottom:9px; }
.color-presets { display:grid;grid-template-columns:1fr 1fr;gap:6px; }
.color-presets button { display:flex;align-items:center;gap:7px;padding:5px 7px;font-size:12px; }
.color-presets i,.region-swatch { display:inline-block;width:13px;height:13px;border-radius:3px;border:1px solid #0002;flex:none; }
.custom-color { margin-top:12px;font-size:12px; }
.color-inputs { display:flex;gap:6px;margin:6px 0; }
.color-inputs input { border:1px solid #c8cbd6;border-radius:4px;height:31px;min-width:0;background:white;color:#424769; }
.color-inputs input[type=color] { width:38px;padding:2px;flex:none;cursor:pointer; }
.color-inputs input[type=text] { width:100%;padding:5px;font:12px monospace; }
.inherit-color { width:100%;font-size:12px; }
.color-controls p { margin:8px 0;font-size:11px; }
.region-list button { display:flex;gap:5px;align-items:center; }
aside h3 { margin:24px 0 10px;font-size:14px;font-weight:600; }
aside p { font-size:12px;color:#73798b;margin:16px 0; }
.region-list { display:flex;gap:6px;flex-wrap:wrap;max-height:175px;overflow:auto; }
.selection-actions { display:flex;gap:6px;margin-top:12px; }
footer { padding:12px 20px;display:flex;align-items:center;gap:10px;border-top:1px solid #e0e1e8; }
footer p { flex:1;margin:0;font-size:12px;color:#70778b; }footer p.error { color:#a43644; }
button.save { background:#514268;border-color:#514268;color:#fff; }
@media(max-width:760px) { .editor-workspace { grid-template-columns:minmax(0,1fr) 165px; }header>span{display:none}.editor-toolbar{padding:8px}.preview-label{margin-left:0} }
</style>
