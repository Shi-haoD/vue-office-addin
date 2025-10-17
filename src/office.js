/* src/office.js */
export async function writeCell(value = 'Hello Vue!') {
	await Excel.run(async (context) => {
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		sheet.getRange('A1').values = [[value]];
		await context.sync();
	});
}

export async function readCell() {
	return await Excel.run(async (context) => {
		const sheet = context.workbook.worksheets.getActiveWorksheet();
		const range = sheet.getRange('A1');
		range.load('values');
		await context.sync();
		return range.values[0][0];
	});
}
