// Build script: bundles every nodes/**/*.node.ts into dist/ as a self-contained
// CommonJS file (exceljs and all other dependencies are inlined, so the published
// package has no runtime dependencies — a requirement for n8n verified nodes).
// n8n-workflow stays external: it is provided by n8n itself at run time.
const esbuild = require('esbuild');
const fs = require('fs');
const path = require('path');

const rootDir = path.join(__dirname, '..');
const nodesDir = path.join(rootDir, 'nodes');
const distDir = path.join(rootDir, 'dist');

function walk(dir) {
	let results = [];
	for (const entry of fs.readdirSync(dir, { withFileTypes: true })) {
		const full = path.join(dir, entry.name);
		if (entry.isDirectory()) {
			results = results.concat(walk(full));
		} else {
			results.push(full);
		}
	}
	return results;
}

const entryPoints = walk(nodesDir).filter((f) => f.endsWith('.node.ts'));
if (entryPoints.length === 0) {
	throw new Error('No *.node.ts entry points found under nodes/');
}

for (const entry of entryPoints) {
	const rel = path.relative(nodesDir, entry);
	const outRel = rel.replace(/\.ts$/, '.js');
	const outfile = path.join(distDir, 'nodes', outRel);
	fs.mkdirSync(path.dirname(outfile), { recursive: true });
	esbuild.buildSync({
		entryPoints: [entry],
		outfile,
		bundle: true,
		platform: 'node',
		format: 'cjs',
		target: 'node18',
		sourcemap: false,
		minify: false,
		legalComments: 'none',
		external: ['n8n-workflow'],
	});
	console.log(`  bundled ${rel} -> dist/nodes/${outRel}`);
}

let copied = 0;
for (const file of walk(nodesDir)) {
	const ext = path.extname(file);
	if (ext === '.json' || ext === '.svg') {
		const rel = path.relative(nodesDir, file);
		const dest = path.join(distDir, 'nodes', rel);
		fs.mkdirSync(path.dirname(dest), { recursive: true });
		fs.copyFileSync(file, dest);
		copied++;
		console.log(`  copied ${rel}`);
	}
}
console.log(`build: ${entryPoints.length} node(s) bundled, ${copied} asset(s) copied`);
