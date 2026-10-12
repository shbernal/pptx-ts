<script setup lang="ts">
// The README's banner, imported from where the README reads it so the two never drift. The
// site switches colour mode with a `.dark` class rather than `prefers-color-scheme`, so a
// `<picture>` would follow the OS and not the toggle; both images render and CSS hides one.
import bannerDark from '../../assets/readme/banner/banner-dark.svg?url'
import bannerLight from '../../assets/readme/banner/banner-light.svg?url'

const alt = 'pptx-ts: write a program, get a PowerPoint file'
</script>

<template>
	<div class="home-banner">
		<img class="home-banner__light" :src="bannerLight" :alt="alt" width="1280" height="320" />
		<img class="home-banner__dark" :src="bannerDark" :alt="alt" width="1280" height="320" />
	</div>
</template>
