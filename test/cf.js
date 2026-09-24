// Conditional formatting tests: generate rules via the bundled node and assert
// on the sheet XML inside the produced xlsx (unzipped via PowerShell).
const assert = require('assert');
const fs = require('fs');
const os = require('os');
const path = require('path');
const { execSync } = require('child_process');
const { ExcelJs } = require('../dist/nodes/ExcelJs/ExcelJs.node.js');

const node = new ExcelJs();
const BASE = { resource: 'conditionalFormatting', operation: 'addRule', sheetName: 'Sheet1', cfRange: 'A1:A5', cfFillColor: 'FFFF0000', binaryPropertyInput: 'data', binaryPropertyOutput: 'data', outputFileName: 't.xlsx' };

function ctx(params, items) {
	return {
		getInputData: () => items,
		// mimic n8n: return the fallback when the parameter was not set on the node
		getNodeParameter: (n, _i, fallback) => (params[n] !== undefined ? params[n] : fallback),
		continueOnFail: () => false,
		helpers: {
			getBinaryDataBuffer: async (i, p) => items[i].binary[p].data,
			prepareBinaryData: async (b) => ({ data: b }),
		},
	};
}

async function sheetXml(params, bin) {
	const out = await node.execute.call(ctx(params, [{ json: {}, binary: { data: { data: bin } } }]));
	const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'cf-'));
	const zipPath = path.join(tmp, 't.zip');
	fs.writeFileSync(zipPath, out[0][0].binary.data.data);
	execSync(`powershell.exe -NoProfile -Command "Expand-Archive -Path '${zipPath}' -DestinationPath '${path.join(tmp, 'x')}' -Force"`);
	const xml = fs.readFileSync(path.join(tmp, 'x', 'xl', 'worksheets', 'sheet1.xml'), 'utf8');
	fs.rmSync(tmp, { recursive: true, force: true });
	return xml;
}

(async () => {
	// base workbook to operate on
	let out = await node.execute.call(ctx({ resource: 'workbook', operation: 'create', binaryPropertyOutput: 'data', outputFileName: 't.xlsx' }, [{}]));
	const bin = out[0][0].binary.data.data;

	// 1. containsText must emit type + SEARCH formula built from the text field
	{
		const xml = await sheetXml({ ...BASE, cfType: 'containsText', cfFormula: 'foo' }, bin);
		const cf = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [''])[0];
		assert.ok(cf.includes('type="containsText"'), 'containsText rule must carry type="containsText"');
		assert.ok(/SEARCH\((&quot;|")foo/.test(cf), 'containsText must generate a SEARCH formula from the text field');
		console.log('1. CF containsText OK ->', cf.slice(0, 120) + '...');
	}

	// 2. between needs TWO formulas (min, max)
	{
		const xml = await sheetXml({ ...BASE, cfType: 'cellIs', cfOperator: 'between', cfFormula: '10', cfSecondValue: '20' }, bin);
		const cf = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [''])[0];
		const formulas = cf.match(/<formula>/g) || [];
		assert.strictEqual(formulas.length, 2, 'between must emit exactly two <formula> entries');
		assert.ok(cf.includes('operator="between"'));
		console.log('2. CF between OK -> two <formula> entries');
	}

	// 3. between without second value must throw
	{
		await assert.rejects(
			() => sheetXml({ ...BASE, cfType: 'cellIs', cfOperator: 'between', cfFormula: '10' }, bin),
			/Second Value/,
		);
		console.log('3. CF between missing max -> rejected OK');
	}

	// 4. numeric "0" stays numeric, not a string
	{
		const xml = await sheetXml({ ...BASE, cfType: 'cellIs', cfOperator: 'greaterThan', cfFormula: '0' }, bin);
		const cf = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [''])[0];
		assert.ok(/<formula>0<\/formula>/.test(cf), 'value 0 must be emitted as numeric formula');
		console.log('4. CF numeric zero OK');
	}

	// 5. plain cellIs still works
	{
		const xml = await sheetXml({ ...BASE, cfType: 'cellIs', cfOperator: 'greaterThan', cfFormula: '100' }, bin);
		const cf = (xml.match(/<conditionalFormatting[\s\S]*?<\/conditionalFormatting>/) || [''])[0];
		assert.ok(cf.includes('type="cellIs"') && cf.includes('operator="greaterThan"') && cf.includes('<formula>100</formula>'));
		console.log('5. CF cellIs greaterThan OK');
	}

	console.log('\nAll CF tests passed.');
})().catch((e) => {
	console.error('CF TEST FAILED:', e.message);
	process.exit(1);
});
