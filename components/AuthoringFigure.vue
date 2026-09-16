<script setup>
import { computed, defineAsyncComponent, ref, watch } from 'vue'
import { useSlideContext } from '@slidev/client'
import { annotations } from 'virtual:slidev-authoring/data'
const props = defineProps({ src:String, alt:String, caption:String })
const { $clicks, $frontmatter, $renderContext } = useSlideContext()
const saved = ref(null), naturalRatio = ref(1), editing = ref(false), savedNotice = ref(false)
const id = computed(() => $frontmatter.contentId)
const entry = computed(() => {
  const candidate = saved.value || annotations[id.value]
  return candidate?.src === props.src ? candidate : null
})
const ratio = computed(() => entry.value?.ratio || naturalRatio.value)
const regions = computed(() => entry.value?.regions || [[]])
const active = computed(() => regions.value[Math.min($clicks.value, regions.value.length - 1)] || [])
const href = computed(() => props.src.startsWith('/') ? import.meta.env.BASE_URL + props.src.slice(1) : props.src)
const canEdit = import.meta.env.DEV ? computed(() => ['slide', 'presenter'].includes($renderContext.value)) : false
const Editor = import.meta.env.DEV ? defineAsyncComponent(() => import('../editor/FigureEditor.vue')) : null
watch(href, url => {
  saved.value = null
  const img = new Image()
  img.onload = () => { if (href.value === url && img.naturalHeight) naturalRatio.value = img.naturalWidth / img.naturalHeight }
  img.src = url
}, { immediate:true })
function finish(entry) { saved.value = entry; savedNotice.value = true; editing.value = false }
</script>
<template>
  <figure class="authoring-figure" :data-figure-id="id">
    <div class="authoring-image-box">
      <svg class="focus-overlay" :viewBox="`-1.5 -1.5 103 ${100 / ratio + 3}`" :data-focus-step="$clicks" :data-focus-style="entry?.focusStyle || 'wash'" role="img" :aria-label="alt">
        <image :href="href" x="0" y="0" width="100" :height="100 / ratio" preserveAspectRatio="none" />
        <rect v-for="(r, i) in active" :key="`${$clicks}-${i}`" class="focus-border" :x="r[0]" :y="r[1] / ratio" :width="r[2]" :height="r[3] / ratio" />
      </svg>
      <button v-if="canEdit" class="figure-edit-button" title="Edit this figure's highlights" @click.stop="editing = true; savedNotice = false">{{ savedNotice ? '저장됨 · 강조 편집' : '그림 강조 편집' }}</button>
    </div>
    <figcaption v-if="caption">{{ caption }}</figcaption>
  </figure>
  <Teleport v-if="editing && Editor" to="body">
    <component :is="Editor" :slide-id="id" :src="src" :href="href" :caption="caption" :ratio="ratio" :initial-step="$clicks" @close="editing = false" @saved="finish" />
  </Teleport>
</template>
<style scoped>
.authoring-figure { margin:0; height:100%; display:flex; flex-direction:column; }
.authoring-image-box { flex:1; min-height:0; width:100%; position:relative; }
figcaption { flex:none; text-align:center; font-size:var(--authoring-caption-size, 15px); color:var(--authoring-caption-color, #656565); padding-top:var(--authoring-caption-gap, 8px); }

.focus-overlay { position:absolute; inset:0; width:100%; height:100%; pointer-events:none; }
.focus-border { vector-effect:non-scaling-stroke; rx:.45; fill:var(--authoring-highlight, #7655AB); fill-opacity:.19; stroke:none; mix-blend-mode:multiply; }
.focus-overlay[data-focus-style="frame"] .focus-border { fill:none; stroke:var(--authoring-highlight, #7655AB); stroke-width:2.2px; stroke-opacity:.9; }
.figure-edit-button { position:absolute; right:0; top:0; padding:6px 10px; border:1px solid #c3bed4; border-radius:5px; background:#fff; color:#514268; font:13px Arial,sans-serif; opacity:0; cursor:pointer; z-index:3; }
.authoring-figure:hover .figure-edit-button, .figure-edit-button:focus-visible { opacity:1; }
@media print { .figure-edit-button { display:none; } }
</style>
