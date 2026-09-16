<script setup>
import { computed } from 'vue'
import { bibliography } from 'virtual:slidev-authoring/data'
const props = defineProps({ ids: { type:Array, default:() => [] } })
const entries = computed(() => props.ids.map(id => {
  const ref = bibliography[id]
  if (!ref) return `[${id}] Reference missing from references.json`
  return `[${id}] ${ref.authors} "${ref.title}." ${ref.venue}.`
}))
</script>
<template>
  <div v-if="entries.length" class="authoring-references">
    <div v-for="entry in entries" :key="entry" class="reference-entry">{{ entry }}</div>
  </div>
</template>
<style scoped>
.authoring-references { position:fixed; left:var(--authoring-reference-inset, 40px); right:var(--authoring-reference-inset, 40px); bottom:var(--authoring-reference-bottom, 20px); color:var(--authoring-reference-color, #858993); font:var(--authoring-reference-size, 11px)/1.25 Arial,sans-serif; text-align:left; }
.reference-entry { white-space:normal; overflow-wrap:normal; }
</style>
