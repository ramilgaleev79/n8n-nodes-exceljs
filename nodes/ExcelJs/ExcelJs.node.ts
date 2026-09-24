import type { IExecuteFunctions, INodeExecutionData } from 'n8n-workflow';
import { INodeType, INodeTypeDescription } from 'n8n-workflow';
import * as ExcelJS from 'exceljs';

type ExcelWorkbook = ExcelJS.Workbook;
type ExcelWorksheet = ExcelJS.Worksheet;

interface ParsedRange {
	startCol: number;
	startRow: number;
	endCol: number;
	// null when the range is written with an open end (e.g. "A1:D" = column D down to the last row)
	endRow: number | null;
}

interface ParsedCell {
	col: number;
	row: number;
}

const CELL_ADDR_RE = /^[A-Za-z]{1,3}\d{1,7}$/;
const RANGE_RE = /^([A-Za-z]{1,3})(\d{1,7}):([A-Za-z]{1,3})(\d{1,7})$/;
const OPEN_RANGE_RE = /^([A-Za-z]{1,3})(\d{1,7}):([A-Za-z]{1,3})$/;

// Excel localizes function names; HYPERLINK() is ГИПЕРССЫЛКА() in Russian Excel.
const HYPERLINK_FORMULA_RE = /^(?:HYPERLINK|ГИПЕРССЫЛКА)\s*\(\s*"([^"]+)"\s*[,;]\s*"([^"]+)"\s*\)$/i;

function colLetterToNumber(letters: string): number {
	let col = 0;
	for (const ch of letters.toUpperCase()) {
		col = col * 26 + (ch.charCodeAt(0) - 64);
	}
	return col;
}

// Column number -> Excel letters (correct for columns >= Z, e.g. 27 -> AA)
function colNumberToLetter(num: number): string {
	let letters = '';
	let n = num;
	while (n > 0) {
		const rem = (n - 1) % 26;
		letters = String.fromCharCode(65 + rem) + letters;
		n = Math.floor((n - 1) / 26);
	}
	return letters;
}

function parseCellRef(ref: string): ParsedCell {
	const m = ref.match(/^([A-Za-z]{1,3})(\d{1,7})$/);
	if (!m) {
		throw new Error(`Invalid cell address: "${ref}". Expected format like A1, B2, etc.`);
	}
	return { col: colLetterToNumber(m[1]), row: parseInt(m[2], 10) };
}

function validateCellAddress(addr: string): void {
	if (!CELL_ADDR_RE.test(addr)) {
		throw new Error(`Invalid cell address: "${addr}". Expected format like A1, B2, AA100, etc.`);
	}
}

function parseRange(rangeStr: string): ParsedRange {
	let m = rangeStr.match(RANGE_RE);
	if (m) {
		const startCol = colLetterToNumber(m[1]);
		const startRow = parseInt(m[2], 10);
		const endCol = colLetterToNumber(m[3]);
		const endRow = parseInt(m[4], 10);
		if (startCol > endCol || startRow > endRow) {
			throw new Error(`Invalid range: "${rangeStr}". The start cell must be above and to the left of the end cell.`);
		}
		return { startCol, startRow, endCol, endRow };
	}
	m = rangeStr.match(OPEN_RANGE_RE);
	if (m) {
		const startCol = colLetterToNumber(m[1]);
		const startRow = parseInt(m[2], 10);
		const endCol = colLetterToNumber(m[3]);
		if (startCol > endCol) {
			throw new Error(`Invalid range: "${rangeStr}". The start column must not be to the right of the end column.`);
		}
		return { startCol, startRow, endCol, endRow: null };
	}
	throw new Error(`Invalid range format: "${rangeStr}". Expected format like A1:D10 or A1:D (down to the last row).`);
}

function requireSheet(workbook: ExcelWorkbook, name: string): ExcelWorksheet {
	const ws = workbook.getWorksheet(name);
	if (!ws) {
		const available: string[] = [];
		workbook.eachSheet((s) => available.push(s.name));
		throw new Error(`Worksheet "${name}" not found in the workbook. ` +
			(available.length
				? `Available worksheets: ${available.join(', ')}`
				: 'The workbook contains no worksheets'));
	}
	return ws;
}

function findBinaryItemIndex(allItems: INodeExecutionData[], propName: string): number {
	for (let idx = 0; idx < allItems.length; idx++) {
		if (allItems[idx].binary && allItems[idx].binary![propName]) {
			return idx;
		}
	}
	return -1;
}

// Display labels for the node subtitle, keyed by resource -> operation -> label.
// Kept in sync with the option lists above; smoke-tested by test/subtitle.js.
const OPERATION_LABELS: Record<string, Record<string, string>> = {
	workbook: { create: 'Create', getInfo: 'Get Info', setProperties: 'Set Properties' },
	worksheet: { add: 'Add', getAll: 'Get All', remove: 'Remove', rename: 'Rename', protect: 'Protect', unprotect: 'Unprotect', freeze: 'Freeze Panes', autoFilter: 'Set Auto Filter', setProperties: 'Set Properties', pageSetup: 'Set Page Setup' },
	cell: { readValue: 'Read Value', readHyperlink: 'Read Hyperlink', writeValue: 'Write Value', writeFormula: 'Write Formula', setStyle: 'Set Style', merge: 'Merge Cells', unmerge: 'Unmerge Cells', addHyperlink: 'Add Hyperlink', setNumFmt: 'Set Number Format' },
	row: { addRow: 'Add Row', addRows: 'Add Rows', getRow: 'Get Row', setHeight: 'Set Height', hide: 'Hide', show: 'Show' },
	column: { setWidth: 'Set Width', hide: 'Hide', show: 'Show' },
	range: { writeData: 'Write Data', readData: 'Read Data', readHyperlinks: 'Read Hyperlinks', setStyle: 'Set Style' },
	image: { add: 'Add Image' },
	conditionalFormatting: { addRule: 'Add Rule' },
};

// "0" must stay numeric: Number(v) || v would turn it back into a string (0 is falsy)
function toNumOrStr(raw: string): number | string {
	const trimmed = raw.trim();
	return trimmed !== '' && !isNaN(Number(trimmed)) ? Number(trimmed) : raw;
}

export class ExcelJs implements INodeType {
	description: INodeTypeDescription = {
		displayName: 'ExcelJS',
		name: 'excelJs',
		icon: 'file:exceljs.svg',
		group: ['transform'],
		version: 1,
		subtitle: `={{ ${JSON.stringify(OPERATION_LABELS)}[$parameter.resource][$parameter.operation] || $parameter.operation }}`,
		description: 'Read, create and manipulate Excel files using ExcelJS',
		documentationUrl: 'https://github.com/ramilgaleev79/n8n-nodes-exceljs#readme',
		defaults: {
			name: 'ExcelJS',
		},
		inputs: ['main'],
		outputs: ['main'],
		credentials: [],
		properties: [
			// ─── Resource ───
			{
				displayName: 'Resource',
				name: 'resource',
				type: 'options',
				noDataExpression: true,
				options: [
					{ name: 'Workbook', value: 'workbook' },
					{ name: 'Worksheet', value: 'worksheet' },
					{ name: 'Cell', value: 'cell' },
					{ name: 'Row', value: 'row' },
					{ name: 'Column', value: 'column' },
					{ name: 'Range', value: 'range' },
					{ name: 'Image', value: 'image' },
					{ name: 'Conditional Formatting', value: 'conditionalFormatting' },
				],
				default: 'workbook',
			},
			// ═══ WORKBOOK OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['workbook'] } },
				options: [
					{ name: 'Create', value: 'create', description: 'Create a new empty workbook', action: 'Create a workbook' },
					{ name: 'Get Info', value: 'getInfo', description: 'Get workbook properties and sheet list', action: 'Get workbook info' },
					{ name: 'Set Properties', value: 'setProperties', description: 'Set workbook properties (author, title, etc.)', action: 'Set workbook properties' },
				],
				default: 'create',
			},
			{
				displayName: 'Creator',
				name: 'wbCreator',
				type: 'string',
				default: '',
				displayOptions: { show: { resource: ['workbook'], operation: ['setProperties'] } },
			},
			{
				displayName: 'Title',
				name: 'wbTitle',
				type: 'string',
				default: '',
				displayOptions: { show: { resource: ['workbook'], operation: ['setProperties'] } },
			},
			{
				displayName: 'Subject',
				name: 'wbSubject',
				type: 'string',
				default: '',
				displayOptions: { show: { resource: ['workbook'], operation: ['setProperties'] } },
			},
			{
				displayName: 'Company',
				name: 'wbCompany',
				type: 'string',
				default: '',
				displayOptions: { show: { resource: ['workbook'], operation: ['setProperties'] } },
			},
			// ═══ WORKSHEET OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['worksheet'] } },
				options: [
					{ name: 'Add', value: 'add', description: 'Add a new worksheet', action: 'Add a worksheet' },
					{ name: 'Get All', value: 'getAll', description: 'List all worksheets', action: 'Get all worksheets' },
					{ name: 'Remove', value: 'remove', description: 'Remove a worksheet', action: 'Remove a worksheet' },
					{ name: 'Rename', value: 'rename', description: 'Rename a worksheet', action: 'Rename a worksheet' },
					{ name: 'Protect', value: 'protect', description: 'Protect a worksheet with password', action: 'Protect a worksheet' },
					{ name: 'Unprotect', value: 'unprotect', description: 'Remove worksheet protection', action: 'Unprotect a worksheet' },
					{ name: 'Freeze Panes', value: 'freeze', description: 'Freeze rows and columns', action: 'Freeze panes' },
					{ name: 'Set Auto Filter', value: 'autoFilter', description: 'Set auto filter on a range', action: 'Set auto filter' },
					{ name: 'Set Properties', value: 'setProperties', description: 'Set sheet tab color, visibility, zoom', action: 'Set sheet properties' },
					{ name: 'Set Page Setup', value: 'pageSetup', description: 'Set page orientation, margins, paper size', action: 'Set page setup' },
				],
				default: 'add',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				description: 'Name of the worksheet',
				displayOptions: {
					show: {
						resource: ['worksheet'],
						operation: ['add', 'remove', 'rename', 'protect', 'unprotect', 'freeze', 'autoFilter', 'setProperties', 'pageSetup'],
					},
				},
			},
			{
				displayName: 'New Name',
				name: 'newSheetName',
				type: 'string',
				required: true,
				default: '',
				displayOptions: { show: { resource: ['worksheet'], operation: ['rename'] } },
			},
			{
				displayName: 'Password',
				name: 'sheetPassword',
				type: 'string',
				typeOptions: { password: true },
				default: '',
				displayOptions: { show: { resource: ['worksheet'], operation: ['protect'] } },
			},
			{
				displayName: 'Freeze Rows',
				name: 'freezeRows',
				type: 'number',
				default: 1,
				description: 'Number of rows to freeze from top',
				displayOptions: { show: { resource: ['worksheet'], operation: ['freeze'] } },
			},
			{
				displayName: 'Freeze Columns',
				name: 'freezeColumns',
				type: 'number',
				default: 0,
				description: 'Number of columns to freeze from left',
				displayOptions: { show: { resource: ['worksheet'], operation: ['freeze'] } },
			},
			{
				displayName: 'Range',
				name: 'autoFilterRange',
				type: 'string',
				required: true,
				default: 'A1:D1',
				placeholder: 'A1:Z1',
				displayOptions: { show: { resource: ['worksheet'], operation: ['autoFilter'] } },
			},
			{
				displayName: 'Tab Color',
				name: 'tabColor',
				type: 'color',
				default: '',
				displayOptions: { show: { resource: ['worksheet'], operation: ['setProperties'] } },
			},
			{
				displayName: 'Visibility',
				name: 'sheetState',
				type: 'options',
				options: [
					{ name: 'Visible', value: 'visible' },
					{ name: 'Hidden', value: 'hidden' },
					{ name: 'Very Hidden', value: 'veryHidden' },
				],
				default: 'visible',
				displayOptions: { show: { resource: ['worksheet'], operation: ['setProperties'] } },
			},
			{
				displayName: 'Zoom',
				name: 'sheetZoom',
				type: 'number',
				default: 100,
				description: 'Zoom percentage (10-400)',
				displayOptions: { show: { resource: ['worksheet'], operation: ['setProperties'] } },
			},
			{
				displayName: 'Orientation',
				name: 'pageOrientation',
				type: 'options',
				options: [
					{ name: 'Portrait', value: 'portrait' },
					{ name: 'Landscape', value: 'landscape' },
				],
				default: 'portrait',
				displayOptions: { show: { resource: ['worksheet'], operation: ['pageSetup'] } },
			},
			{
				displayName: 'Paper Size',
				name: 'paperSize',
				type: 'options',
				options: [
					{ name: 'Letter', value: 1 },
					{ name: 'Legal', value: 5 },
					{ name: 'A3', value: 8 },
					{ name: 'A4', value: 9 },
					{ name: 'A5', value: 11 },
				],
				default: 9,
				displayOptions: { show: { resource: ['worksheet'], operation: ['pageSetup'] } },
			},
			{
				displayName: 'Fit to Page',
				name: 'fitToPage',
				type: 'boolean',
				default: false,
				displayOptions: { show: { resource: ['worksheet'], operation: ['pageSetup'] } },
			},
			// ═══ CELL OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['cell'] } },
				options: [
					{ name: 'Read Value', value: 'readValue', description: 'Read a cell value', action: 'Read cell value' },
					{ name: 'Read Hyperlink', value: 'readHyperlink', description: 'Read a hyperlink from a cell', action: 'Read cell hyperlink' },
					{ name: 'Write Value', value: 'writeValue', description: 'Write a value to a cell', action: 'Write cell value' },
					{ name: 'Write Formula', value: 'writeFormula', description: 'Write a formula to a cell', action: 'Write formula' },
					{ name: 'Set Style', value: 'setStyle', description: 'Set cell formatting (font, fill, border, alignment)', action: 'Set cell style' },
					{ name: 'Merge Cells', value: 'merge', description: 'Merge a range of cells', action: 'Merge cells' },
					{ name: 'Unmerge Cells', value: 'unmerge', description: 'Unmerge a range of cells', action: 'Unmerge cells' },
					{ name: 'Add Hyperlink', value: 'addHyperlink', description: 'Add a hyperlink to a cell', action: 'Add hyperlink' },
					{ name: 'Set Number Format', value: 'setNumFmt', description: 'Set number/date format', action: 'Set number format' },
				],
				default: 'writeValue',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				displayOptions: { show: { resource: ['cell'] } },
			},
			{
				displayName: 'Cell',
				name: 'cellAddress',
				type: 'string',
				required: true,
				default: 'A1',
				placeholder: 'A1',
				description: 'Cell address (e.g. A1, B2, C3)',
				displayOptions: {
					show: {
						resource: ['cell'],
						operation: ['readValue', 'readHyperlink', 'writeValue', 'writeFormula', 'setStyle', 'addHyperlink', 'setNumFmt'],
					},
				},
			},
			{
				displayName: 'Value',
				name: 'cellValue',
				type: 'string',
				default: '',
				displayOptions: { show: { resource: ['cell'], operation: ['writeValue'] } },
			},
			{
				displayName: 'Value Type',
				name: 'cellValueType',
				type: 'options',
				options: [
					{ name: 'String', value: 'string' },
					{ name: 'Number', value: 'number' },
					{ name: 'Boolean', value: 'boolean' },
					{ name: 'Date', value: 'date' },
				],
				default: 'string',
				displayOptions: { show: { resource: ['cell'], operation: ['writeValue'] } },
			},
			{
				displayName: 'Formula',
				name: 'cellFormula',
				type: 'string',
				required: true,
				default: '',
				placeholder: 'SUM(A1:A10)',
				description: 'Formula without leading =',
				displayOptions: { show: { resource: ['cell'], operation: ['writeFormula'] } },
			},
			{
				displayName: 'Range',
				name: 'mergeRange',
				type: 'string',
				required: true,
				default: 'A1:D1',
				placeholder: 'A1:D1',
				displayOptions: { show: { resource: ['cell'], operation: ['merge', 'unmerge'] } },
			},
			{
				displayName: 'URL',
				name: 'hyperlinkUrl',
				type: 'string',
				required: true,
				default: '',
				placeholder: 'https://example.com',
				displayOptions: { show: { resource: ['cell'], operation: ['addHyperlink'] } },
			},
			{
				displayName: 'Display Text',
				name: 'hyperlinkText',
				type: 'string',
				default: '',
				displayOptions: { show: { resource: ['cell'], operation: ['addHyperlink'] } },
			},
			{
				displayName: 'Format',
				name: 'numFmt',
				type: 'string',
				required: true,
				default: '#,##0.00',
				placeholder: '#,##0.00 or dd.mm.yyyy or 0.00%',
				displayOptions: { show: { resource: ['cell'], operation: ['setNumFmt'] } },
			},
			{
				displayName: 'Font Bold',
				name: 'fontBold',
				type: 'boolean',
				default: false,
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Font Italic',
				name: 'fontItalic',
				type: 'boolean',
				default: false,
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Font Size',
				name: 'fontSize',
				type: 'number',
				default: 0,
				description: 'Font size (0 = keep default)',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Font Name',
				name: 'fontName',
				type: 'string',
				default: '',
				placeholder: 'Arial',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Font Color (ARGB)',
				name: 'fontColor',
				type: 'string',
				default: '',
				placeholder: 'FF0000FF',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Fill Color (ARGB)',
				name: 'fillColor',
				type: 'string',
				default: '',
				placeholder: 'FFFFFF00',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Horizontal Alignment',
				name: 'hAlign',
				type: 'options',
				options: [
					{ name: 'None', value: '' },
					{ name: 'Left', value: 'left' },
					{ name: 'Center', value: 'center' },
					{ name: 'Right', value: 'right' },
					{ name: 'Fill', value: 'fill' },
					{ name: 'Justify', value: 'justify' },
				],
				default: '',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Vertical Alignment',
				name: 'vAlign',
				type: 'options',
				options: [
					{ name: 'None', value: '' },
					{ name: 'Top', value: 'top' },
					{ name: 'Middle', value: 'middle' },
					{ name: 'Bottom', value: 'bottom' },
				],
				default: '',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Wrap Text',
				name: 'wrapText',
				type: 'boolean',
				default: false,
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Border Style',
				name: 'borderStyle',
				type: 'options',
				options: [
					{ name: 'None', value: '' },
					{ name: 'Thin', value: 'thin' },
					{ name: 'Medium', value: 'medium' },
					{ name: 'Thick', value: 'thick' },
					{ name: 'Double', value: 'double' },
					{ name: 'Dotted', value: 'dotted' },
					{ name: 'Dashed', value: 'dashed' },
				],
				default: '',
				description: 'Apply this border style to all sides',
				displayOptions: { show: { resource: ['cell'], operation: ['setStyle'] } },
			},
			// ═══ ROW OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['row'] } },
				options: [
					{ name: 'Add Row', value: 'addRow', description: 'Add a single row of data', action: 'Add a row' },
					{ name: 'Add Rows', value: 'addRows', description: 'Add multiple rows from input items', action: 'Add rows' },
					{ name: 'Get Row', value: 'getRow', description: 'Read a row by number', action: 'Get a row' },
					{ name: 'Set Height', value: 'setHeight', description: 'Set row height', action: 'Set row height' },
					{ name: 'Hide', value: 'hide', description: 'Hide a row', action: 'Hide a row' },
					{ name: 'Show', value: 'show', description: 'Show a hidden row', action: 'Show a row' },
				],
				default: 'addRow',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				displayOptions: { show: { resource: ['row'] } },
			},
			{
				displayName: 'Row Number',
				name: 'rowNumber',
				type: 'number',
				required: true,
				default: 1,
				displayOptions: {
					show: { resource: ['row'], operation: ['getRow', 'setHeight', 'hide', 'show'] },
				},
			},
			{
				displayName: 'Values (JSON Array)',
				name: 'rowValues',
				type: 'string',
				required: true,
				default: '["value1", "value2", "value3"]',
				description: 'JSON array of values for the row',
				displayOptions: { show: { resource: ['row'], operation: ['addRow'] } },
			},
			{
				displayName: 'Fields to Use',
				name: 'rowFields',
				type: 'string',
				default: '',
				placeholder: 'name, email, phone',
				description: 'Comma-separated field names from input items. Empty = use all fields.',
				displayOptions: { show: { resource: ['row'], operation: ['addRows'] } },
			},
			{
				displayName: 'Add Header Row',
				name: 'addHeader',
				type: 'boolean',
				default: true,
				description: 'Whether to add a header row with field names',
				displayOptions: { show: { resource: ['row'], operation: ['addRows'] } },
			},
			{
				displayName: 'Start Cell',
				name: 'startCell',
				type: 'string',
				required: true,
				default: 'A1',
				description: 'Top-left cell to start writing data',
				displayOptions: { show: { resource: ['row'], operation: ['addRows'] } },
			},
			{
				displayName: 'Height',
				name: 'rowHeight',
				type: 'number',
				required: true,
				default: 20,
				displayOptions: { show: { resource: ['row'], operation: ['setHeight'] } },
			},
			// ═══ COLUMN OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['column'] } },
				options: [
					{ name: 'Set Width', value: 'setWidth', description: 'Set column width', action: 'Set column width' },
					{ name: 'Hide', value: 'hide', description: 'Hide a column', action: 'Hide a column' },
					{ name: 'Show', value: 'show', description: 'Show a hidden column', action: 'Show a column' },
				],
				default: 'setWidth',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				displayOptions: { show: { resource: ['column'] } },
			},
			{
				displayName: 'Column',
				name: 'columnKey',
				type: 'string',
				required: true,
				default: 'A',
				placeholder: 'A or 1',
				displayOptions: { show: { resource: ['column'] } },
			},
			{
				displayName: 'Width',
				name: 'columnWidth',
				type: 'number',
				required: true,
				default: 20,
				displayOptions: { show: { resource: ['column'], operation: ['setWidth'] } },
			},
			// ═══ RANGE OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['range'] } },
				options: [
					{ name: 'Write Data', value: 'writeData', description: 'Write input items to a range', action: 'Write data to range' },
					{ name: 'Read Data', value: 'readData', description: 'Read data from a range to JSON', action: 'Read data from range' },
					{ name: 'Read Hyperlinks', value: 'readHyperlinks', description: 'Read hyperlinks from a range of cells', action: 'Read hyperlinks from range' },
					{ name: 'Set Style', value: 'setStyle', description: 'Apply style to a range of cells', action: 'Set range style' },
				],
				default: 'writeData',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				displayOptions: { show: { resource: ['range'] } },
			},
			{
				displayName: 'Range',
				name: 'rangeAddress',
				type: 'string',
				required: true,
				default: 'A1:D10',
				placeholder: 'A1:D10 or A1:D (down to the last row)',
				displayOptions: { show: { resource: ['range'] } },
			},
			{
				displayName: 'First Row Is Header',
				name: 'headerRow',
				type: 'boolean',
				default: true,
				displayOptions: { show: { resource: ['range'], operation: ['readData'] } },
			},
			{
				displayName: 'Fields',
				name: 'rangeFields',
				type: 'string',
				default: '',
				placeholder: 'name, email, phone',
				description: 'Comma-separated field names. Empty = use all fields from input items.',
				displayOptions: { show: { resource: ['range'], operation: ['writeData'] } },
			},
			{
				displayName: 'Add Header Row',
				name: 'rangeAddHeader',
				type: 'boolean',
				default: true,
				displayOptions: { show: { resource: ['range'], operation: ['writeData'] } },
			},
			// Range > Set Style
			{
				displayName: 'Font Bold',
				name: 'fontBold',
				type: 'boolean',
				default: false,
				displayOptions: { show: { resource: ['range'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Fill Color (ARGB)',
				name: 'fillColor',
				type: 'string',
				default: '',
				placeholder: 'FFFFFF00',
				displayOptions: { show: { resource: ['range'], operation: ['setStyle'] } },
			},
			{
				displayName: 'Border Style',
				name: 'borderStyle',
				type: 'options',
				options: [
					{ name: 'None', value: '' },
					{ name: 'Thin', value: 'thin' },
					{ name: 'Medium', value: 'medium' },
					{ name: 'Thick', value: 'thick' },
				],
				default: '',
				displayOptions: { show: { resource: ['range'], operation: ['setStyle'] } },
			},
			// ═══ IMAGE OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['image'] } },
				options: [
					{ name: 'Add Image', value: 'add', description: 'Add an image to a worksheet', action: 'Add an image' },
				],
				default: 'add',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				displayOptions: { show: { resource: ['image'] } },
			},
			{
				displayName: 'Binary Property',
				name: 'imageBinaryProperty',
				type: 'string',
				required: true,
				default: 'data',
				description: 'Name of the binary property containing the image',
				displayOptions: { show: { resource: ['image'], operation: ['add'] } },
			},
			{
				displayName: 'Image Extension',
				name: 'imageExtension',
				type: 'options',
				options: [
					{ name: 'PNG', value: 'png' },
					{ name: 'JPEG', value: 'jpeg' },
					{ name: 'GIF', value: 'gif' },
				],
				default: 'png',
				displayOptions: { show: { resource: ['image'], operation: ['add'] } },
			},
			{
				displayName: 'Top-Left Cell (Col)',
				name: 'imgTlCol',
				type: 'number',
				default: 0,
				displayOptions: { show: { resource: ['image'], operation: ['add'] } },
			},
			{
				displayName: 'Top-Left Cell (Row)',
				name: 'imgTlRow',
				type: 'number',
				default: 0,
				displayOptions: { show: { resource: ['image'], operation: ['add'] } },
			},
			{
				displayName: 'Bottom-Right Cell (Col)',
				name: 'imgBrCol',
				type: 'number',
				default: 5,
				displayOptions: { show: { resource: ['image'], operation: ['add'] } },
			},
			{
				displayName: 'Bottom-Right Cell (Row)',
				name: 'imgBrRow',
				type: 'number',
				default: 5,
				displayOptions: { show: { resource: ['image'], operation: ['add'] } },
			},
			// ═══ CONDITIONAL FORMATTING OPERATIONS ═══
			{
				displayName: 'Operation',
				name: 'operation',
				type: 'options',
				noDataExpression: true,
				displayOptions: { show: { resource: ['conditionalFormatting'] } },
				options: [
					{ name: 'Add Rule', value: 'addRule', description: 'Add a conditional formatting rule', action: 'Add a CF rule' },
				],
				default: 'addRule',
			},
			{
				displayName: 'Sheet Name',
				name: 'sheetName',
				type: 'string',
				required: true,
				default: 'Sheet1',
				displayOptions: { show: { resource: ['conditionalFormatting'] } },
			},
			{
				displayName: 'Range',
				name: 'cfRange',
				type: 'string',
				required: true,
				default: 'A1:A100',
				displayOptions: { show: { resource: ['conditionalFormatting'], operation: ['addRule'] } },
			},
			{
				displayName: 'Rule Type',
				name: 'cfType',
				type: 'options',
				options: [
					{ name: 'Cell Is', value: 'cellIs' },
					{ name: 'Contains Text', value: 'containsText' },
				],
				default: 'cellIs',
				displayOptions: { show: { resource: ['conditionalFormatting'], operation: ['addRule'] } },
			},
			{
				displayName: 'Operator',
				name: 'cfOperator',
				type: 'options',
				options: [
					{ name: 'Greater Than', value: 'greaterThan' },
					{ name: 'Less Than', value: 'lessThan' },
					{ name: 'Equal', value: 'equal' },
					{ name: 'Between', value: 'between' },
				],
				default: 'greaterThan',
				displayOptions: { show: { resource: ['conditionalFormatting'], operation: ['addRule'], cfType: ['cellIs'] } },
			},
			{
				displayName: 'Value / Text',
				name: 'cfFormula',
				type: 'string',
				required: true,
				default: '100',
				description: 'Comparison value, or the text to search for when the rule type is Contains Text. For Between this is the minimum value.',
				displayOptions: { show: { resource: ['conditionalFormatting'], operation: ['addRule'] } },
			},
			{
				displayName: 'Second Value (Max)',
				name: 'cfSecondValue',
				type: 'string',
				required: true,
				default: '',
				placeholder: '200',
				description: 'Maximum value, only used with the Between operator',
				displayOptions: { show: { resource: ['conditionalFormatting'], operation: ['addRule'], cfOperator: ['between'] } },
			},
			{
				displayName: 'Fill Color (ARGB)',
				name: 'cfFillColor',
				type: 'string',
				required: true,
				default: 'FF00FF00',
				displayOptions: { show: { resource: ['conditionalFormatting'], operation: ['addRule'] } },
			},
			// ═══ COMMON: Binary input field ═══
			{
				displayName: 'Binary Property (Input)',
				name: 'binaryPropertyInput',
				type: 'string',
				required: true,
				default: 'data',
				description: 'Name of binary property containing the input Excel file',
				displayOptions: {
					hide: {
						resource: ['workbook'],
						operation: ['create'],
					},
				},
			},
			{
				displayName: 'Output Binary Property',
				name: 'binaryPropertyOutput',
				type: 'string',
				required: true,
				default: 'data',
				description: 'Name of binary property to write the output Excel file to',
			},
			{
				displayName: 'Output File Name',
				name: 'outputFileName',
				type: 'string',
				required: true,
				default: 'output.xlsx',
			},
		],
	};

	async execute(this: IExecuteFunctions): Promise<INodeExecutionData[][]> {
		const items = this.getInputData();
		const returnData: INodeExecutionData[] = [];
		const resource = this.getNodeParameter('resource', 0) as string;
		const operation = this.getNodeParameter('operation', 0) as string;

		// ─── Batch operations: writeData, addRows ───
		// These need binary from ONE item but JSON data from ALL items.
		// Process them once and return, instead of per-item loop.
		const isBatchWrite = (resource === 'range' && operation === 'writeData') ||
			(resource === 'row' && operation === 'addRows');

		if (isBatchWrite) {
			// Wrapped in try/catch so continueOnFail works the same as the per-item path
			try {
				if (items.length === 0) {
					throw new Error('No input items to write. ' +
						'Make sure the previous node passes at least one item.');
				}
				const inputProp = this.getNodeParameter('binaryPropertyInput', 0, 'data') as string;
				const outputProp = this.getNodeParameter('binaryPropertyOutput', 0, 'data') as string;
				const outputName = this.getNodeParameter('outputFileName', 0, 'output.xlsx') as string;

				const binaryIdx = findBinaryItemIndex(items, inputProp);
				if (binaryIdx === -1) {
					throw new Error(`None of the ${items.length} input items contains binary data in property "${inputProp}". ` +
						'Connect a node that passes an Excel file (e.g. Read Binary File, HTTP Request). ' +
						'Use Merge (Append) to combine the data and the file into a single stream.');
				}
				let binaryBuf: Buffer;
				try {
					binaryBuf = await this.helpers.getBinaryDataBuffer(binaryIdx, inputProp);
				} catch (e: any) {
					throw new Error(`Failed to read binary data from property "${inputProp}" (item ${binaryIdx}): ${e.message}`);
				}
				if (!binaryBuf || binaryBuf.length === 0) {
					throw new Error(`Binary data in property "${inputProp}" is empty. ` +
						'Make sure the previous node passes a non-empty Excel file.');
				}

				const workbook = new ExcelJS.Workbook();
				try {
					await workbook.xlsx.load(binaryBuf);
				} catch (e: any) {
					const msg = e.message || '';
					throw new Error(`Failed to load the Excel file: ${msg}. ` +
						'Check that the file is a valid .xlsx document.');
				}

				const sheetName = this.getNodeParameter('sheetName', 0) as string;
				const ws = requireSheet(workbook, sheetName);

				const dataItems = items.filter(item => Object.keys(item.json).length > 0);
				if (dataItems.length === 0) {
					throw new Error('All input items have empty JSON data. Nothing to write.');
				}

				let fields: string[];
				let fieldsStr = '';
				let addHeader = true;
				let ref: ParsedCell;

				if (resource === 'range' && operation === 'writeData') {
					const rangeAddr = this.getNodeParameter('rangeAddress', 0) as string;
					fieldsStr = this.getNodeParameter('rangeFields', 0, '') as string;
					addHeader = this.getNodeParameter('rangeAddHeader', 0, true) as boolean;
					const rangeStart = rangeAddr.split(':')[0];
					ref = parseCellRef(rangeStart);
				} else {
					fieldsStr = this.getNodeParameter('rowFields', 0, '') as string;
					addHeader = this.getNodeParameter('addHeader', 0, true) as boolean;
					const startCell = this.getNodeParameter('startCell', 0, 'A1') as string;
					ref = parseCellRef(startCell);
				}

				if (fieldsStr.trim()) {
					fields = fieldsStr.split(',').map(f => f.trim()).filter(f => f);
					if (fields.length === 0) {
						throw new Error('The field list is empty. Specify fields separated by commas.');
					}
				} else {
					fields = Object.keys(dataItems[0].json);
					if (fields.length === 0) {
						throw new Error('The input item contains no JSON fields to write.');
					}
				}

				let currentRow = ref.row;
				const debugRows: any[] = [];
				if (addHeader) {
					const row = ws.getRow(currentRow);
					fields.forEach((f, idx) => {
						row.getCell(ref.col + idx).value = f;
					});
					row.commit();
					currentRow++;
				}
				for (let di = 0; di < dataItems.length; di++) {
					const item = dataItems[di];
					const row = ws.getRow(currentRow);
					const rowDebug: any = { rowNum: currentRow, values: {} };
					fields.forEach((f, idx) => {
						const val = item.json[f] ?? '';
						const cell = row.getCell(ref.col + idx);
						if (typeof val === 'string' && val.startsWith('=')) {
							const formulaBody = val.substring(1);
							const hyperlinkMatch = formulaBody.match(HYPERLINK_FORMULA_RE);
							if (hyperlinkMatch) {
								cell.value = {
									text: hyperlinkMatch[2],
									hyperlink: hyperlinkMatch[1],
								};
								cell.font = { color: { argb: 'FF0563C1' }, underline: true };
							} else {
								cell.value = { formula: formulaBody };
							}
						} else {
							cell.value = val as any;
						}
						rowDebug.values[f] = val;
					});
					row.commit();
					debugRows.push(rowDebug);
					currentRow++;
				}

				let buffer: Buffer;
				try {
					buffer = Buffer.from(await workbook.xlsx.writeBuffer());
				} catch (e: any) {
					throw new Error(`Failed to save the Excel file: ${e.message}. The workbook may contain invalid data.`);
				}
				const binaryData = await this.helpers.prepareBinaryData(buffer, outputName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
				returnData.push({
					json: {
						success: true,
						totalItems: items.length,
						dataItemsCount: dataItems.length,
						fields,
						rowsWritten: dataItems.length,
						debug: debugRows,
					},
					binary: { [outputProp]: binaryData },
				});
				return [returnData];
			} catch (error: any) {
				if (this.continueOnFail()) {
					returnData.push({
						json: {
							error: error.message || 'Unknown error',
							resource,
							operation,
						},
					});
					return [returnData];
				}
				throw error;
			}
		}

		// ─── Per-item operations (all other operations) ───
		for (let i = 0; i < items.length; i++) {
			try {
				const workbook = new ExcelJS.Workbook();
				const outputProp = this.getNodeParameter('binaryPropertyOutput', i, 'data') as string;
				const outputName = this.getNodeParameter('outputFileName', i, 'output.xlsx') as string;

				// ─── Load existing workbook (unless creating new) ───
				if (!(resource === 'workbook' && operation === 'create')) {
					const inputProp = this.getNodeParameter('binaryPropertyInput', i, 'data') as string;
					if (!items[i].binary) {
						throw new Error(`Input data (item ${i}) contains no binary data. ` +
							'Connect a node that passes an Excel file (e.g. Read Binary File or HTTP Request).');
					}
					if (!items[i].binary![inputProp]) {
						const availableKeys = Object.keys(items[i].binary!);
						throw new Error(`Binary property "${inputProp}" not found in item ${i}. ` +
							(availableKeys.length
								? `Available properties: ${availableKeys.join(', ')}`
								: 'The item contains no binary properties'));
					}
					let binaryData: Buffer;
					try {
						binaryData = await this.helpers.getBinaryDataBuffer(i, inputProp);
					} catch (e: any) {
						throw new Error(`Failed to read binary data from property "${inputProp}": ${e.message}`);
					}
					if (!binaryData || binaryData.length === 0) {
						throw new Error(`Binary data in property "${inputProp}" is empty. ` +
							'Make sure the previous node passes a non-empty Excel file.');
					}
					try {
						await workbook.xlsx.load(binaryData);
					} catch (e: any) {
						const msg = e.message || '';
						if (msg.includes('End of data') || msg.includes('Format') || msg.includes('Unexpected')) {
							throw new Error('Failed to open the file: the file is corrupted or in an unsupported format. ' +
								'Make sure a valid .xlsx file is passed (not .xls, .csv or another format). ' +
								`Original error: ${msg}`);
						}
						throw new Error(`Failed to load the Excel file: ${msg}. Check that the file is a valid .xlsx document.`);
					}
				}

				let jsonOutput: any = {};

				// Read-only operations do not serialize an output file (same as range/readData):
				// workbook.getInfo, worksheet.getAll, cell.readValue, cell.readHyperlink, row.getRow.
				const isReadOnly =
					(resource === 'workbook' && operation === 'getInfo') ||
					(resource === 'worksheet' && operation === 'getAll') ||
					(resource === 'cell' && (operation === 'readValue' || operation === 'readHyperlink')) ||
					(resource === 'row' && operation === 'getRow');

				// ═══ WORKBOOK ═══
				if (resource === 'workbook') {
					if (operation === 'create') {
						workbook.addWorksheet('Sheet1');
					} else if (operation === 'getInfo') {
						const sheets: any[] = [];
						workbook.eachSheet((ws: ExcelWorksheet) => {
							sheets.push({
								id: ws.id,
								name: ws.name,
								state: ws.state,
								rowCount: ws.rowCount,
								columnCount: ws.columnCount,
							});
						});
						jsonOutput = {
							creator: workbook.creator || '',
							title: workbook.title || '',
							subject: workbook.subject || '',
							company: workbook.company || '',
							sheetCount: sheets.length,
							sheets,
						};
					} else if (operation === 'setProperties') {
						const creator = this.getNodeParameter('wbCreator', i, '') as string;
						const title = this.getNodeParameter('wbTitle', i, '') as string;
						const subject = this.getNodeParameter('wbSubject', i, '') as string;
						const company = this.getNodeParameter('wbCompany', i, '') as string;
						if (creator) workbook.creator = creator;
						if (title) workbook.title = title;
						if (subject) workbook.subject = subject;
						if (company) workbook.company = company;
					}
				}

				// ═══ WORKSHEET ═══
				if (resource === 'worksheet') {
					const sheetName = this.getNodeParameter('sheetName', i, 'Sheet1') as string;
					if (operation === 'add') {
						if (workbook.getWorksheet(sheetName)) {
							throw new Error(`Worksheet "${sheetName}" already exists in the workbook. Choose another name or remove the existing sheet.`);
						}
						workbook.addWorksheet(sheetName);
					} else if (operation === 'getAll') {
						const sheets: any[] = [];
						workbook.eachSheet((ws: ExcelWorksheet) => {
							sheets.push({ id: ws.id, name: ws.name, state: ws.state });
						});
						jsonOutput = { sheets };
					} else if (operation === 'remove') {
						const ws = requireSheet(workbook, sheetName);
						workbook.removeWorksheet(ws.id);
					} else if (operation === 'rename') {
						const newName = this.getNodeParameter('newSheetName', i) as string;
						if (!newName.trim()) {
							throw new Error('The new worksheet name cannot be empty.');
						}
						if (workbook.getWorksheet(newName)) {
							throw new Error(`A worksheet named "${newName}" already exists.`);
						}
						const ws = requireSheet(workbook, sheetName);
						ws.name = newName;
					} else if (operation === 'protect') {
						const password = this.getNodeParameter('sheetPassword', i, '') as string;
						const ws = requireSheet(workbook, sheetName);
						await ws.protect(password, {});
					} else if (operation === 'unprotect') {
						const ws = requireSheet(workbook, sheetName);
						ws.unprotect();
					} else if (operation === 'freeze') {
						const rows = this.getNodeParameter('freezeRows', i, 1) as number;
						const cols = this.getNodeParameter('freezeColumns', i, 0) as number;
						if (rows < 0 || cols < 0) {
							throw new Error('The number of frozen rows/columns cannot be negative.');
						}
						const ws = requireSheet(workbook, sheetName);
						// Correct topLeftCell for columns >= Z (colNumberToLetter instead of fromCharCode);
						// spread the existing view so zoom and other settings survive
						ws.views = [{
							...((ws.views[0] as any) || {}),
							state: 'frozen',
							xSplit: cols,
							ySplit: rows,
							topLeftCell: `${colNumberToLetter(cols + 1)}${rows + 1}`,
							activeCell: 'A1',
						}];
					} else if (operation === 'autoFilter') {
						const range = this.getNodeParameter('autoFilterRange', i) as string;
						if (!RANGE_RE.test(range)) {
							throw new Error(`Invalid range format for auto filter: "${range}". Expected format A1:D1.`);
						}
						const ws = requireSheet(workbook, sheetName);
						ws.autoFilter = range;
					} else if (operation === 'setProperties') {
						const ws = requireSheet(workbook, sheetName);
						const tabColor = this.getNodeParameter('tabColor', i, '') as string;
						const state = this.getNodeParameter('sheetState', i, 'visible') as string;
						const zoom = this.getNodeParameter('sheetZoom', i, 100) as number;
						if (zoom < 10 || zoom > 400) {
							throw new Error(`Zoom must be between 10 and 400. Got: ${zoom}.`);
						}
						if (tabColor) ws.properties.tabColor = { argb: tabColor.replace('#', '') };
						ws.state = state as any;
						// spread the existing view so frozen panes survive setting the zoom
						ws.views = [{ ...((ws.views[0] as any) || {}), zoomScale: zoom }];
					} else if (operation === 'pageSetup') {
						const ws = requireSheet(workbook, sheetName);
						const orientation = this.getNodeParameter('pageOrientation', i) as string;
						const paperSize = this.getNodeParameter('paperSize', i) as number;
						const fitToPage = this.getNodeParameter('fitToPage', i) as boolean;
						ws.pageSetup = {
							...ws.pageSetup,
							orientation: orientation as any,
							paperSize,
							fitToPage,
						};
					}
				}

				// ═══ CELL ═══
				if (resource === 'cell') {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const ws = requireSheet(workbook, sheetName);
					if (operation === 'readValue') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const cell = ws.getCell(addr);
						jsonOutput = {
							address: addr,
							value: cell.value,
							type: cell.type,
							formula: cell.formula || null,
							text: cell.text || '',
						};
					} else if (operation === 'readHyperlink') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const cell = ws.getCell(addr);
						const v = cell.value;
						jsonOutput = {
							address: addr,
							hyperlink: cell.hyperlink || null,
							text: (v && typeof v === 'object' && (v as any).text !== undefined) ? String((v as any).text) : (cell.text != null ? String(cell.text) : ''),
						};
					} else if (operation === 'writeValue') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const rawValue = this.getNodeParameter('cellValue', i) as string;
						const valueType = this.getNodeParameter('cellValueType', i) as string;
						const cell = ws.getCell(addr);
						if (valueType === 'number') {
							const num = Number(rawValue);
							if (isNaN(num)) {
								throw new Error(`Cannot convert "${rawValue}" to a number for cell ${addr}. ` +
									'Make sure the value is a valid number (use a dot as the decimal separator: 3.14).');
							}
							cell.value = num;
						} else if (valueType === 'boolean') {
							const lower = rawValue.toLowerCase().trim();
							if (lower !== 'true' && lower !== 'false') {
								throw new Error(`Cannot convert "${rawValue}" to a boolean for cell ${addr}. Valid values: true, false.`);
							}
							cell.value = lower === 'true';
						} else if (valueType === 'date') {
							const date = new Date(rawValue);
							if (isNaN(date.getTime())) {
								throw new Error(`Cannot convert "${rawValue}" to a date for cell ${addr}. Use ISO format: 2026-02-18 or 2026-02-18T10:30:00.`);
							}
							cell.value = date;
						} else {
							cell.value = rawValue;
						}
					} else if (operation === 'writeFormula') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const formula = this.getNodeParameter('cellFormula', i) as string;
						if (!formula.trim()) {
							throw new Error('The formula cannot be empty.');
						}
						ws.getCell(addr).value = { formula };
					} else if (operation === 'setStyle') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const cell = ws.getCell(addr);
						const bold = this.getNodeParameter('fontBold', i) as boolean;
						const italic = this.getNodeParameter('fontItalic', i) as boolean;
						const size = this.getNodeParameter('fontSize', i) as number;
						const fontName = this.getNodeParameter('fontName', i, '') as string;
						const fontColor = this.getNodeParameter('fontColor', i, '') as string;
						const fillColor = this.getNodeParameter('fillColor', i, '') as string;
						const hAlign = this.getNodeParameter('hAlign', i, '') as string;
						const vAlign = this.getNodeParameter('vAlign', i, '') as string;
						const wrapText = this.getNodeParameter('wrapText', i) as boolean;
						const borderStyle = this.getNodeParameter('borderStyle', i, '') as string;
						if (size < 0) {
							throw new Error(`Font size cannot be negative: ${size}.`);
						}
						const font: any = {};
						if (bold) font.bold = true;
						if (italic) font.italic = true;
						if (size > 0) font.size = size;
						if (fontName) font.name = fontName;
						if (fontColor) font.color = { argb: fontColor };
						if (Object.keys(font).length) cell.font = font;
						if (fillColor) {
							cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
						}
						const alignment: any = {};
						if (hAlign) alignment.horizontal = hAlign;
						if (vAlign) alignment.vertical = vAlign;
						if (wrapText) alignment.wrapText = true;
						if (Object.keys(alignment).length) cell.alignment = alignment;
						if (borderStyle) {
							const side = { style: borderStyle as any };
							cell.border = { top: side, left: side, bottom: side, right: side };
						}
					} else if (operation === 'merge') {
						const range = this.getNodeParameter('mergeRange', i) as string;
						parseRange(range);
						ws.mergeCells(range);
					} else if (operation === 'unmerge') {
						const range = this.getNodeParameter('mergeRange', i) as string;
						parseRange(range);
						ws.unMergeCells(range);
					} else if (operation === 'addHyperlink') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const url = this.getNodeParameter('hyperlinkUrl', i) as string;
						if (!url.trim()) {
							throw new Error('The hyperlink URL cannot be empty.');
						}
						const text = this.getNodeParameter('hyperlinkText', i, '') as string;
						const cell = ws.getCell(addr);
						cell.value = { text: text || url, hyperlink: url };
					} else if (operation === 'setNumFmt') {
						const addr = this.getNodeParameter('cellAddress', i) as string;
						validateCellAddress(addr);
						const fmt = this.getNodeParameter('numFmt', i) as string;
						if (!fmt.trim()) {
							throw new Error('The number format cannot be empty. Examples: #,##0.00, dd.mm.yyyy, 0.00%');
						}
						ws.getCell(addr).numFmt = fmt;
					}
				}

				// ═══ ROW ═══
				if (resource === 'row') {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const ws = requireSheet(workbook, sheetName);
					// addRows is handled by the batch path above, not here
					if (operation === 'addRow') {
						const raw = this.getNodeParameter('rowValues', i) as string;
						let values: any[];
						try {
							values = JSON.parse(raw);
						} catch (e: any) {
							throw new Error(`Failed to parse JSON in the "Values" field: ${e.message}. Expected a JSON array, e.g. ["value1", "value2", 123]`);
						}
						if (!Array.isArray(values)) {
							throw new Error('The "Values" field must contain a JSON array (starting with [ and ending with ]). ' +
								`Got: ${typeof values}`);
						}
						ws.addRow(values);
					} else if (operation === 'getRow') {
						const rowNum = this.getNodeParameter('rowNumber', i) as number;
						if (rowNum < 1) {
							throw new Error(`Row number must be >= 1. Got: ${rowNum}.`);
						}
						const row = ws.getRow(rowNum);
						const values: any[] = [];
						row.eachCell({ includeEmpty: true }, (cell) => {
							values.push(cell.value);
						});
						jsonOutput = { row: rowNum, values };
					} else if (operation === 'setHeight') {
						const rowNum = this.getNodeParameter('rowNumber', i) as number;
						const height = this.getNodeParameter('rowHeight', i) as number;
						if (rowNum < 1) {
							throw new Error(`Row number must be >= 1. Got: ${rowNum}.`);
						}
						if (height <= 0) {
							throw new Error(`Row height must be > 0. Got: ${height}.`);
						}
						ws.getRow(rowNum).height = height;
					} else if (operation === 'hide') {
						const rowNum = this.getNodeParameter('rowNumber', i) as number;
						if (rowNum < 1) throw new Error(`Row number must be >= 1. Got: ${rowNum}.`);
						ws.getRow(rowNum).hidden = true;
					} else if (operation === 'show') {
						const rowNum = this.getNodeParameter('rowNumber', i) as number;
						if (rowNum < 1) throw new Error(`Row number must be >= 1. Got: ${rowNum}.`);
						ws.getRow(rowNum).hidden = false;
					}
				}

				// ═══ COLUMN ═══
				if (resource === 'column') {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const ws = requireSheet(workbook, sheetName);
					const colKey = this.getNodeParameter('columnKey', i) as string;
					if (!colKey.trim()) {
						throw new Error('The column identifier cannot be empty. Specify a letter (A, B, C) or a number (1, 2, 3).');
					}
					const col = ws.getColumn(colKey);
					if (operation === 'setWidth') {
						const width = this.getNodeParameter('columnWidth', i) as number;
						if (width <= 0) {
							throw new Error(`Column width must be > 0. Got: ${width}.`);
						}
						col.width = width;
					} else if (operation === 'hide') {
						col.hidden = true;
					} else if (operation === 'show') {
						col.hidden = false;
					}
				}

				// ═══ RANGE ═══
				if (resource === 'range') {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const ws = requireSheet(workbook, sheetName);
					const rangeAddr = this.getNodeParameter('rangeAddress', i) as string;
					// writeData is handled by the batch path above, not here
					if (operation === 'readData') {
						const headerRow = this.getNodeParameter('headerRow', i, true) as boolean;
						const { startCol, startRow, endCol, endRow } = parseRange(rangeAddr);
						const endRowFinal = endRow === null ? ws.rowCount : endRow;
						let headers: string[] = [];
						let dataStartRow = startRow;
						if (headerRow) {
							for (let c = startCol; c <= endCol; c++) {
								headers.push(String(ws.getCell(startRow, c).value ?? ''));
							}
							dataStartRow = startRow + 1;
						} else {
							for (let c = startCol; c <= endCol; c++) {
								headers.push(`col_${c}`);
							}
						}
						const readItems: INodeExecutionData[] = [];
						for (let r = dataStartRow; r <= endRowFinal; r++) {
							const obj: any = {};
							for (let c = startCol; c <= endCol; c++) {
								obj[headers[c - startCol]] = ws.getCell(r, c).value;
							}
							readItems.push({ json: obj });
						}
						returnData.push(...readItems);
						continue; // read-only: skip output file serialization
					} else if (operation === 'readHyperlinks') {
						const { startCol, startRow, endCol, endRow } = parseRange(rangeAddr);
						const endRowFinal = endRow === null ? ws.rowCount : endRow;
						const readItems: INodeExecutionData[] = [];
						for (let r = startRow; r <= endRowFinal; r++) {
							for (let c = startCol; c <= endCol; c++) {
								const cell = ws.getCell(r, c);
								const v = cell.value;
								const text = (v && typeof v === 'object' && (v as any).text !== undefined) ? String((v as any).text) : (cell.text != null ? String(cell.text) : '');
								readItems.push({
									json: {
										address: colNumberToLetter(c) + r,
										hyperlink: cell.hyperlink || null,
										text,
									},
								});
							}
						}
						returnData.push(...readItems);
						continue; // read-only: skip output file serialization
					} else if (operation === 'setStyle') {
						const bold = this.getNodeParameter('fontBold', i, false) as boolean;
						const fillColor = this.getNodeParameter('fillColor', i, '') as string;
						const borderStyle = this.getNodeParameter('borderStyle', i, '') as string;
						const { startCol, startRow, endCol, endRow } = parseRange(rangeAddr);
						if (endRow === null) {
							throw new Error('For the setStyle operation, specify a full range (e.g. A1:D10).');
						}
						for (let r = startRow; r <= endRow; r++) {
							for (let c = startCol; c <= endCol; c++) {
								const cell = ws.getCell(r, c);
								if (bold) cell.font = { ...cell.font, bold: true };
								if (fillColor) {
									cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } };
								}
								if (borderStyle) {
									const side = { style: borderStyle as any };
									cell.border = { top: side, left: side, bottom: side, right: side };
								}
							}
						}
					}
				}

				// ═══ IMAGE ═══
				if (resource === 'image' && operation === 'add') {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const ws = requireSheet(workbook, sheetName);
					const imgProp = this.getNodeParameter('imageBinaryProperty', i) as string;
					const ext = this.getNodeParameter('imageExtension', i) as string;
					const tlCol = this.getNodeParameter('imgTlCol', i) as number;
					const tlRow = this.getNodeParameter('imgTlRow', i) as number;
					const brCol = this.getNodeParameter('imgBrCol', i) as number;
					const brRow = this.getNodeParameter('imgBrRow', i) as number;
					if (tlCol >= brCol || tlRow >= brRow) {
						throw new Error(`Invalid image coordinates: the top-left corner (${tlCol},${tlRow}) must be above and to the left of the bottom-right corner (${brCol},${brRow}).`);
					}
					if (!items[i].binary || !items[i].binary![imgProp]) {
						throw new Error(`Binary property "${imgProp}" with the image not found in item ${i}. Connect a node that passes an image (Read Binary File, HTTP Request, etc.).`);
					}
					let imgBuffer: Buffer;
					try {
						imgBuffer = await this.helpers.getBinaryDataBuffer(i, imgProp);
					} catch (e: any) {
						throw new Error(`Failed to read the image from "${imgProp}": ${e.message}`);
					}
					if (!imgBuffer || imgBuffer.length === 0) {
						throw new Error(`The image in property "${imgProp}" is empty (0 bytes).`);
					}
					const imageId = workbook.addImage({ buffer: imgBuffer, extension: ext as any });
					ws.addImage(imageId, {
						tl: { col: tlCol, row: tlRow } as any,
						br: { col: brCol, row: brRow } as any,
					});
				}

				// ═══ CONDITIONAL FORMATTING ═══
				if (resource === 'conditionalFormatting' && operation === 'addRule') {
					const sheetName = this.getNodeParameter('sheetName', i) as string;
					const ws = requireSheet(workbook, sheetName);
					const range = this.getNodeParameter('cfRange', i) as string;
					const parsedCfRange = parseRange(range);
					if (parsedCfRange.endRow === null) {
						throw new Error('For conditional formatting, specify a full range (e.g. A1:A10).');
					}
					const cfType = this.getNodeParameter('cfType', i) as string;
					const formulaVal = this.getNodeParameter('cfFormula', i) as string;
					const fillColor = this.getNodeParameter('cfFillColor', i) as string;
					if (!formulaVal.trim()) {
						throw new Error('The value/text for conditional formatting cannot be empty.');
					}
					// For a solid pattern the fill color is set via fgColor (not bgColor)
					const style = { fill: { type: 'pattern', pattern: 'solid', fgColor: { argb: fillColor } } };
					let rule: any;
					if (cfType === 'containsText') {
						// ExcelJS expects text + operator for containsText (formulae is ignored)
						rule = { type: 'containsText', operator: 'containsText', text: formulaVal, style };
					} else {
						const formulae: (number | string)[] = [toNumOrStr(formulaVal)];
						const cfOperator = this.getNodeParameter('cfOperator', i) as string;
						if (cfOperator === 'between') {
							const secondRaw = this.getNodeParameter('cfSecondValue', i, '') as string;
							if (!String(secondRaw).trim()) {
								throw new Error('The "Second Value (Max)" is required for the Between operator.');
							}
							formulae.push(toNumOrStr(String(secondRaw)));
						}
						rule = { type: 'cellIs', operator: cfOperator, formulae, style };
					}
					ws.addConditionalFormatting({ ref: range, rules: [rule] });
				}

				// Read-only operations do not serialize an output file
				if (isReadOnly) {
					returnData.push({ json: jsonOutput });
					continue;
				}

				// ═══ OUTPUT: Save workbook to binary ═══
				let buffer: Buffer;
				try {
					buffer = Buffer.from(await workbook.xlsx.writeBuffer());
				} catch (e: any) {
					throw new Error(`Failed to save the Excel file: ${e.message}. The workbook may contain invalid data.`);
				}
				const binaryData = await this.helpers.prepareBinaryData(buffer, outputName, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
				returnData.push({
					json: jsonOutput,
					binary: { [outputProp]: binaryData },
				});
			} catch (error: any) {
				if (this.continueOnFail()) {
					const errMsg = error.message || 'Unknown error';
					returnData.push({
						json: {
							error: errMsg,
							resource,
							operation,
							itemIndex: i,
						},
					});
					continue;
				}
				throw error;
			}
		}
		return [returnData];
	}
}
