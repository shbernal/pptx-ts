<script setup>
import { computed, onBeforeUnmount, onMounted, ref } from 'vue'
import { counted, failureMessage, slideList, summarizeNotes } from '../demos/deck-preview.ts'
import SlideFrame from '../demos/SlideFrame.vue'
import { downloadSnippet, previewSnippet } from './live-example.ts'

// What a `ts live` fence renders as, after its code. `load` imports the fence's compiled
// body; `www/live/fence.ts` hoists that import into the page.
const props = defineProps({
	load: { type: Function, required: true },
	fileName: { type: String, required: true },
})

// Nothing runs until the example scrolls near the viewport: a page with eight examples
// should not build eight decks for a reader who stops at the first.
const preview = ref({ status: 'waiting', deck: null, error: '' })
const download = ref({ status: 'idle', error: '' })
const host = ref(null)
let observer = null
// The page is pre-rendered, so the button exists before it does anything. Disabled until
// mount, for the reason `DeckPreview.vue` gives.
const ready = ref(false)

const differences = computed(() => (preview.value.deck ? summarizeNotes(preview.value.deck.fidelity) : []))
const aspect = computed(() => ({ '--deck-aspect': String(preview.value.deck?.aspectRatio ?? 16 / 9) }))

async function render() {
	preview.value = { status: 'rendering', deck: null, error: '' }
	try {
		preview.value = { status: 'ready', deck: await previewSnippet(props.load), error: '' }
	} catch (error) {
		preview.value = { status: 'failed', deck: null, error: failureMessage(error) }
	}
}

async function save() {
	download.value = { status: 'saving', error: '' }
	try {
		await downloadSnippet(props.load, props.fileName)
		download.value = { status: 'saved', error: '' }
	} catch (error) {
		download.value = { status: 'failed', error: failureMessage(error) }
	}
}

onMounted(() => {
	ready.value = true
	observer = new IntersectionObserver(
		(entries) => {
			if (!entries.some((entry) => entry.isIntersecting)) return
			observer?.disconnect()
			observer = null
			render()
		},
		{ rootMargin: '200px' }
	)
	observer.observe(host.value)
})
onBeforeUnmount(() => observer?.disconnect())
</script>

<template>
	<figure ref="host" class="live-example" :style="aspect" :data-status="preview.status">
		<div class="live-example__canvas">
			<template v-if="preview.deck">
				<SlideFrame
					v-for="slide in preview.deck.slides"
					:key="slide.number"
					:markup="slide.markup"
					:styles="preview.deck.styles"
				/>
			</template>
			<div v-else-if="preview.status === 'failed'" class="live-example__placeholder" role="alert">
				<span>This example could not be rendered: {{ preview.error }}</span>
				<button type="button" class="live-example__button" @click="render">Try again</button>
			</div>
			<div v-else class="live-example__placeholder">
				<span>{{ preview.status === 'rendering' ? 'Building the slide…' : 'The slide this code builds' }}</span>
			</div>
		</div>
		<figcaption class="live-example__bar">
			<span class="live-example__label">
				Built in this tab from the code above<template v-if="preview.deck && preview.deck.slides.length > 1">
					· {{ counted(preview.deck.slides.length, 'slide') }}</template
				>
			</span>
			<button type="button" class="live-example__button" :disabled="!ready || download.status === 'saving'" @click="save">
				{{ download.status === 'saving' ? 'Building…' : 'Download .pptx' }}
			</button>
		</figcaption>
		<p v-if="download.status === 'failed'" role="alert" class="live-example__note live-example__note--bad">
			{{ download.error }}
		</p>
		<details v-if="differences.length" class="live-example__note">
			<summary>The preview renderer could not carry back {{ counted(differences.length, 'construct') }}</summary>
			<ul>
				<li v-for="note in differences" :key="note.key">
					<code>{{ note.construct }}</code> {{ note.disposition }}
					<template v-if="preview.deck.slides.length > 1">
						(slide{{ note.slides.length === 1 ? '' : 's' }} {{ slideList(note.slides) }})</template
					>: {{ note.detail }}
				</li>
			</ul>
			<p>The download is the real output; open it in PowerPoint to see these as built.</p>
		</details>
	</figure>
</template>
