// Smoke test: runs the bundled node without n8n using a mocked execution context.
// Covers: workbook create, cell write, addRows batch, open-range readData,
// addHyperlink/readHyperlink, range readHyperlinks.
const assert = require('assert');
const ExcelJS = require('exceljs');
const { ExcelJs } = require('../dist/nodes/ExcelJs/ExcelJs.node.js');

function makeCtx(params, items) {
	return {
		getInputData: () => items,
		getNodeParameter: (name) => params[name],
		continueOnFail: () => false,
		helpers: {
			getBinaryDataBuffer: async (i, prop) => items[i].binary[prop].data,
			prepareBinaryData: async (buf, fileName) => ({ data: buf, fileName, mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' }),
		},
	};
}

(async () => {
	const node = new ExcelJs();
	let bin;

	// 1. workbook:create
	{
		const out = await node.execute.call(makeCtx({ resource: 'workbook', operation: 'create', binaryPropertyOutput: 'data', outputFileName: 't.xlsx' }, [{}]));
		bin = out[0][0].binary.data; // IBinaryData-like mock object; the raw buffer is in .data
		assert.ok(bin.data && bin.data.length > 0, 'create produced a file');
		console.log('1. workbook:create OK,', bin.length, 'bytes');
	}

	// 2. row:addRows batch write
	{
		const items = [
			{ json: {}, binary: { data: bin } },
			{ json: { name: 'Alice', score: 5, link: '=HYPERLINK("https://example.com", "Example")' } },
			{ json: { name: 'Bob', score: 7, link: 'https://n8n.io' } },
		];
		const ctx = makeCtx({ resource: 'row', operation: 'addRows', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx', sheetName: 'Sheet1', rowFields: '', addHeader: true, startCell: 'A1' }, items);
		const out = await node.execute.call(ctx);
		bin = out[0][0].binary.data;
		assert.equal(out[0][0].json.rowsWritten, 2, 'two data rows written');
		console.log('2. row:addRows OK, rowsWritten =', out[0][0].json.rowsWritten);
	}

	// 3. range:readData with open range "A1:C" (ported 0.2.0 patch)
	{
		const items = [{ json: {}, binary: { data: bin } }];
		const ctx = makeCtx({ resource: 'range', operation: 'readData', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx', sheetName: 'Sheet1', rangeAddress: 'A1:C', headerRow: true }, items);
		const out = await node.execute.call(ctx);
		assert.equal(out[0].length, 2, 'open range read two rows');
		assert.equal(out[0][0].json.name, 'Alice');
		assert.equal(out[0][1].json.score, 7);
		console.log('3. range:readData open range OK,', out[0].length, 'rows');
	}

	// 4. cell:addHyperlink + cell:readHyperlink (ported 0.2.0 patch)
	{
		let items = [{ json: {}, binary: { data: bin } }];
		let ctx = makeCtx({ resource: 'cell', operation: 'addHyperlink', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx', sheetName: 'Sheet1', cellAddress: 'E1', hyperlinkUrl: 'https://docs.n8n.io', hyperlinkText: 'Docs' }, items);
		let out = await node.execute.call(ctx);
		bin = out[0][0].binary.data;

		items = [{ json: {}, binary: { data: bin } }];
		ctx = makeCtx({ resource: 'cell', operation: 'readHyperlink', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx', sheetName: 'Sheet1', cellAddress: 'E1' }, items);
		out = await node.execute.call(ctx);
		assert.equal(out[0][0].json.hyperlink, 'https://docs.n8n.io');
		assert.equal(out[0][0].json.text, 'Docs');
		console.log('4. cell:readHyperlink OK ->', out[0][0].json.hyperlink);
	}

	// 5. range:readHyperlinks (ported 0.2.0 patch)
	{
		const items = [{ json: {}, binary: { data: bin } }];
		const ctx = makeCtx({ resource: 'range', operation: 'readHyperlinks', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx', sheetName: 'Sheet1', rangeAddress: 'C2:C' }, items);
		const out = await node.execute.call(ctx);
		const found = out[0].find((it) => it.json.hyperlink === 'https://example.com');
		assert.ok(found, 'hyperlink from addRows found by readHyperlinks');
		assert.equal(found.json.text, 'Example');
		console.log('5. range:readHyperlinks OK ->', found.json.address);
	}

	// 6. range:setStyle must reject open range (guard)
	{
		const items = [{ json: {}, binary: { data: bin } }];
		const ctx = makeCtx({ resource: 'range', operation: 'setStyle', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx', sheetName: 'Sheet1', rangeAddress: 'A1:C', fontBold: true, fillColor: '', borderStyle: '' }, items);
		await assert.rejects(() => node.execute.call(ctx), /full range/);
		console.log('6. range:setStyle open-range guard OK');
	}

	console.log('\nAll smoke tests passed.');
})().catch((e) => {
	console.error('SMOKE TEST FAILED:', e.message);
	process.exit(1);
});
