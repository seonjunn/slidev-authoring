import test from 'node:test'
import assert from 'node:assert/strict'
import { mkdtemp, readFile, writeFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { createAnnotationStore } from '../lib/server.mjs'
import { citationIds, authoringMarkdown } from '../lib/markdown.mjs'
import MarkdownExit from 'markdown-exit'

async function fixture(t) {
  const root = await mkdtemp(join(tmpdir(), 'seminar-editor-'))
  t.after(() => rm(root, {recursive:true, force:true}))
  const source = '---\ncontentId: one\nclicks: 1\n---\n# A\n\n![A](/figures/a.png)\n\n---\ncontentId: two\nclicks: 1\n---\n# B\n\n![B](/figures/b.png)\n'
  await writeFile(join(root, 'slides.md'), source)
  const entry = src => ({src,ratio:2,focusStyle:'wash',regions:[[],[[10,20,30,40]]]})
  await writeFile(join(root, 'annotations.json'), JSON.stringify({one:entry('/figures/a.png'),two:entry('/figures/b.png')}))
  const store = createAnnotationStore(root)
  const payload = async id => {const src=`/figures/${id==='one'?'a':'b'}.png`, data=await store.get(id,src);return{slideId:id,src,revision:data.revision,...data.entry}}
  return {root,store,payload,source}
}

test('saving annotations preserves Markdown, other slides, and untouched click steps', async t => {
  const {root,store,payload,source} = await fixture(t)
  const p = await payload('one'), previous = await store.get('two','/figures/b.png')
  p.regions[1][0][0] = 15
  await store.save(p)
  assert.equal((await store.get('one',p.src)).entry.regions[1][0][0],15)
  assert.deepEqual((await store.get('one',p.src)).entry.regions[0],[])
  assert.deepEqual(await store.get('two','/figures/b.png'),previous)
  assert.equal(await readFile(join(root,'slides.md'),'utf8'),source)
})

test('stale editors cannot overwrite a more recent save', async t => {
  const {store,payload} = await fixture(t)
  const first=await payload('one'), stale=await payload('one')
  first.regions[1][0][0]=12; await store.save(first)
  stale.regions[1][0][0]=35
  await assert.rejects(store.save(stale),e=>e.status===409)
  assert.equal((await store.get('one',first.src)).entry.regions[1][0][0],12)
})

test('concurrent saves to distinct slides are both retained', async t => {
  const {store,payload}=await fixture(t)
  const one=await payload('one'),two=await payload('two')
  one.regions[1][0][0]=11;two.regions[1][0][0]=22
  await Promise.all([store.save(one),store.save(two)])
  assert.equal((await store.get('one',one.src)).entry.regions[1][0][0],11)
  assert.equal((await store.get('two',two.src)).entry.regions[1][0][0],22)
})

test('unknown slides, changed images, invalid geometry and click mismatch cannot modify the file', async t => {
  const {root,store,payload}=await fixture(t)
  const initial=await readFile(join(root,'annotations.json'),'utf8'), p=await payload('one')
  for(const bad of [ {...p,slideId:'../../outside'}, {...p,src:'/figures/replacement.png'}, {...p,regions:[[],[[1,2,-3,4]]]}, {...p,regions:[[]]}, {...p,regions:[[],[[0,0,null,4]]]} ]) await assert.rejects(store.save(bad))
  assert.equal(await readFile(join(root,'annotations.json'),'utf8'),initial)
})

test('ordinary Markdown images, tables and citations render without authored HTML', async () => {
  const md=new MarkdownExit({html:true});authoringMarkdown(md)
  const result=await md.renderAsync('# Title\n\n- A claim [1, 2]\n\n![Figure caption](/figures/a.png)\n\n| A | B |\n| --- | --- |\n| One | Two |')
  assert.match(result,/<li>A claim <sup class="authoring-cite">\[1,2\]<\/sup>/)
  assert.match(result,/<AuthoringFigure src="\/figures\/a.png"/)
  assert.doesNotMatch(result,/<p><div class="authoring-visual">/)
  assert.match(result,/<table>/)
  assert.match(result,/<AuthoringReferences :ids='\[1,2\]' \/>/)
  assert.deepEqual(citationIds('A [2] and [1, 2]; [3](https://example.org); `array[4]`'),[1,2])
})

test('figure and per-region colors survive saves without changing geometry or other slides', async t => {
  const {store,payload,root,source}=await fixture(t)
  const p=await payload('one'), two=await store.get('two','/figures/b.png')
  p.focusColor='#397DA8'; p.regions[1][0].push('#21877E')
  const saved=await store.save(p)
  assert.equal(saved.entry.focusColor,'#397DA8')
  assert.deepEqual(saved.entry.regions,[[],[[10,20,30,40,'#21877E']]])
  assert.deepEqual(await store.get('two','/figures/b.png'),two)
  assert.equal(await readFile(join(root,'slides.md'),'utf8'),source)
  const reset=await payload('one'); delete reset.focusColor; reset.regions[1][0].pop()
  const original=await store.save(reset)
  assert.equal(Object.hasOwn(original.entry,'focusColor'),false)
  assert.deepEqual(original.entry.regions,[[],[[10,20,30,40]]])
})

test('invalid colors cannot modify annotations', async t => {
  const {store,payload,root}=await fixture(t), p=await payload('one')
  const initial=await readFile(join(root,'annotations.json'),'utf8')
  for (const color of ['red','#abc','#12345678','url(https://example.org)',null,12]) {
    await assert.rejects(store.save({...p,focusColor:color}),e=>e.status===400)
    await assert.rejects(store.save({...p,regions:[[],[[10,20,30,40,color]]]}),e=>e.status===400)
  }
  assert.equal(await readFile(join(root,'annotations.json'),'utf8'),initial)
})

test('display colors fall back from region to figure to theme', async () => {
  const {highlightStyle}=await import('../lib/colors.mjs')
  assert.deepEqual(highlightStyle([1,2,3,4]),{})
  assert.deepEqual(highlightStyle([1,2,3,4],'#397DA8'),{'--region-highlight':'#397DA8'})
  assert.deepEqual(highlightStyle([1,2,3,4,'#21877E'],'#397DA8'),{'--region-highlight':'#21877E'})
  assert.deepEqual(highlightStyle([1,2,3,4,'invalid'],'invalid'),{})
})
