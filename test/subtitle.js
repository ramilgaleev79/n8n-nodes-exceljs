// Subtitle test: every resource/operation pair declared in the node properties
// must resolve to a human-readable label through the subtitle expression.
const assert = require('assert');
const { ExcelJs } = require('../dist/nodes/ExcelJs/ExcelJs.node.js');

const node = new ExcelJs();
const subtitle = node.description.subtitle;
assert.ok(subtitle.startsWith('={{') && subtitle.endsWith('}}'), 'subtitle must be an expression');
const body = subtitle.slice(3, -2);
const resolve = new Function('$parameter', `return ${body};`);

const props = node.description.properties;
const resources = props.find((p) => p.name === 'resource').options.map((o) => o.value);

let checked = 0;
for (const prop of props) {
	if (prop.name !== 'operation') continue;
	const resource = prop.displayOptions.show.resource[0];
	for (const op of prop.options) {
		const label = resolve({ resource, operation: op.value });
		assert.strictEqual(label, op.name, `${resource}.${op.value}: subtitle "${label}" != option name "${op.name}"`);
		checked++;
	}
}
console.log(`Subtitle test OK: ${checked} operations across ${resources.length} resources all resolve to their display names.`);
