<template>
	<div class="hello-excel">
		<h2>Vue2 + Office Add-in</h2>
		<button @click="writeData">写入 Excel</button>
		<button @click="readData">读取 Excel</button>
		<p v-if="cellValue">A1 内容：{{ cellValue }}</p>
	</div>
</template>

<script>
import { writeCell, readCell } from '../office';

export default {
	data() {
		return { cellValue: '' };
	},
	methods: {
		async writeData() {
			await writeCell('Hello from Vue!');
		},
		async readData() {
			this.cellValue = await readCell();
		},
	},
};
</script>

<style scoped>
.hello-excel {
	padding: 20px;
	font-family: Arial;
}
button {
	margin-right: 10px;
	padding: 6px 12px;
	cursor: pointer;
}
</style>
