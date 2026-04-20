import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const formulaSpecs = [
	{
		key: "text_multisection_accounting",
		row: 2,
		formula: '=TEXT(-1234.567,"[Green]$#,##0.00;[Red]($#,##0.00);0.00")',
	},
	{
		key: "text_conditional_percent_sections",
		row: 3,
		formula: '=TEXT(0.375,"[>=1]0.0%;[Red](0.0%);0.0%")',
	},
	{ key: "ref_3d_sum", row: 4, formula: "=SUM(Jan:Mar!B2)" },
	{ key: "offset_sum", row: 5, formula: "=SUM(OFFSET(Data!B2,0,0,4,1))" },
	{ key: "indirect_sum", row: 6, formula: '=SUM(INDIRECT(Config!B1&"!"&Config!B2))' },
	{ key: "map_double_sum", row: 7, formula: "=SUM(MAP(Data!B2:B5,LAMBDA(x,x*2)))" },
	{ key: "reduce_square_sum", row: 8, formula: "=REDUCE(0,Data!B2:B5,LAMBDA(a,x,a+x^2))" },
	{ key: "let_sumproduct", row: 9, formula: "=LET(q,Data!B2:B5,p,Data!C2:C5,SUM(q*p))" },
	{ key: "xlookup_price_c", row: 10, formula: '=XLOOKUP("C",Data!A2:A5,Data!C2:C5)' },
	{
		key: "index_xmatch_price_d",
		row: 11,
		formula: '=INDEX(Data!C2:C5,XMATCH("D",Data!A2:A5))',
	},
	{
		key: "textjoin_map",
		row: 12,
		formula:
			'=TEXTJOIN(",",TRUE,MAP(Data!A2:A5,Data!B2:B5,LAMBDA(item,qty,item&":"&qty)))',
	},
	{
		key: "sumproduct_filtered_qty",
		row: 13,
		formula: "=SUMPRODUCT((Data!C2:C5>15)*Data!B2:B5)",
	},
	{ key: "sequence_sum", row: 14, formula: "=SUM(SEQUENCE(3))" },
	{
		key: "byrow_sum",
		row: 15,
		formula: "=SUM(BYROW(Data!B2:C5,LAMBDA(r,SUM(r))))",
	},
	{
		key: "choosecols_sum_qty",
		row: 16,
		formula: "=SUM(CHOOSECOLS(Data!B2:C5,1))",
	},
	{ key: "take_sum_first2", row: 17, formula: "=SUM(TAKE(Data!B2:B5,2))" },
	{ key: "drop_sum_last2", row: 18, formula: "=SUM(DROP(Data!B2:B5,2))" },
];

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const fixturesRoot = path.join(repoRoot, "fixtures");
const outputRoot = path.join(repoRoot, "outputs");

await fs.mkdir(outputRoot, { recursive: true });

function cellValue(ws, address) {
	return ws.getRange(address).values?.[0]?.[0] ?? null;
}

function cellFormula(ws, address) {
	return ws.getRange(address).formulas?.[0]?.[0] ?? null;
}

function captureSummary(summary) {
	return formulaSpecs.map((spec) => ({
		key: spec.key,
		label: cellValue(summary, `A${spec.row}`),
		expected: cellValue(summary, `B${spec.row}`),
		formula: cellFormula(summary, `C${spec.row}`),
		value: cellValue(summary, `C${spec.row}`),
	}));
}

async function importWorkbook(fileName) {
	const input = await FileBlob.load(path.join(fixturesRoot, fileName));
	return SpreadsheetFile.importXlsx(input);
}

async function exportWorkbook(workbook, fileName) {
	const outputPath = path.join(outputRoot, fileName);
	const output = await SpreadsheetFile.exportXlsx(workbook);
	await output.save(outputPath);
	return outputPath;
}

async function runCase(name, fn) {
	try {
		const result = await fn();
		return { name, status: "ok", ...result };
	} catch (error) {
		return {
			name,
			status: "error",
			error: error?.message ?? String(error),
			stack: error?.stack?.split("\n").slice(0, 8) ?? [],
		};
	}
}

const results = [];

results.push(
	await runCase("case10", async () => {
		const workbook = await importWorkbook("case10_formula_fixture.xlsx");
		const summary = workbook.worksheets.getItem("Summary");

		const before = captureSummary(summary);
		summary.getRange("C2:C18").formulas = formulaSpecs.map((spec) => [spec.formula]);
		workbook.recalculate();
		const after = captureSummary(summary);
		const outputPath = await exportWorkbook(workbook, "case10_codex.xlsx");

		return {
			fixture: "fixtures/case10_formula_fixture.xlsx",
			outputPath,
			before,
			after,
		};
	}),
);

const outputPath = path.join(outputRoot, "formula_correctness_codex_results.json");
await fs.writeFile(
	outputPath,
	`${JSON.stringify(
		{
			generatedAt: new Date().toISOString(),
			repoRoot,
			results,
		},
		null,
		2,
	)}\n`,
);

console.log(JSON.stringify({ outputPath, results }, null, 2));
