import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const fixturesRoot = path.join(repoRoot, "fixtures");
const outputRoot = path.join(repoRoot, "outputs");

await fs.mkdir(outputRoot, { recursive: true });

function matrixValues(ws, address) {
	return ws.getRange(address).values ?? [];
}

function cellFormula(ws, address) {
	return ws.getRange(address).formulas?.[0]?.[0] ?? null;
}

function captureCase8(summary, workbook) {
	return {
		sheets: workbook.worksheets.items.map((sheet) => sheet.name),
		qtyFormula: cellFormula(summary, "B2"),
		totalFormula: cellFormula(summary, "B3"),
		spillAnchorFormula: cellFormula(summary, "D2"),
		spillValues: matrixValues(summary, "D2:F4"),
	}
}

function captureCase9(summary, data) {
	return {
		qtyFormula: cellFormula(summary, "B2"),
		priceFormula: cellFormula(summary, "B3"),
		spillAnchorFormula: cellFormula(summary, "E2"),
		spillValues: matrixValues(summary, "E2:J6"),
		dataValues: matrixValues(data, "A1:D6"),
	}
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

const results = []

results.push(
	await runCase("case8", async () => {
		const workbook = await importWorkbook("case8_rename_fixture.xlsx");
		const summary = workbook.worksheets.getItem("Summary");

		const before = captureCase8(summary, workbook);
		const renameResult = workbook.apply([
			{ op: "sheet.set", target: "Data", props: { name: "Renamed" } },
		]);
		workbook.recalculate();
		const after = captureCase8(summary, workbook);
		const outputPath = await exportWorkbook(workbook, "case8_codex.xlsx");

		return {
			fixture: "fixtures/case8_rename_fixture.xlsx",
			outputPath,
			renameResult,
			before,
			after,
			expected: {
				qtyFormula: "=SUM(Renamed!B2:B4)",
				totalFormula: "=SUMPRODUCT(Renamed!B2:B4, Renamed!C2:C4)",
				spillAnchorFormula: "=TRANSPOSE(Renamed!A2:C4)",
				spillValues: [
					["A", "B", "C"],
					[1, 2, 3],
					[10, 20, 30],
				],
			},
		};
	}),
);

results.push(
	await runCase("case9", async () => {
		const workbook = await importWorkbook("case9_shift_fixture.xlsx");
		const summary = workbook.worksheets.getItem("Summary");
		const data = workbook.worksheets.getItem("Data");

		const before = captureCase9(summary, data);
		const operations = [
			{ op: "rows.insert", target: { sheet: "Data", range: "4:4" }, props: { count: 1 } },
			{ op: "columns.insert", target: { sheet: "Data", range: "C:C" }, props: { count: 1 } },
			{ op: "rows.delete", target: { sheet: "Data", range: "5:5" } },
			{ op: "columns.delete", target: { sheet: "Data", range: "C:C" } },
		];
		const applyResults = operations.map((operation) => ({
			op: operation.op,
			result: workbook.apply([operation]),
		}));
		workbook.recalculate();
		const after = captureCase9(summary, data);
		const outputPath = await exportWorkbook(workbook, "case9_codex.xlsx");

		return {
			fixture: "fixtures/case9_shift_fixture.xlsx",
			outputPath,
			applyResults,
			before,
			after,
			expectedFinal: {
				qtyFormula: "=SUM(Data!B2:B5)",
				priceFormula: "=SUM(Data!C2:C5)",
				spillAnchorFormula: "=TRANSPOSE(Data!A2:C5)",
				dataValues: [
					["Item", "Qty", "Price", null],
					["A", 1, 10, null],
					["B", 2, 20, null],
					["X", 10, 100, null],
					["D", 4, 40, null],
					[null, null, null, null],
				],
				spillValues: [
					["A", "B", "X", "D", null, null],
					[1, 2, 10, 4, null, null],
					[10, 20, 100, 40, null, null],
					[null, null, null, null, null, null],
					[null, null, null, null, null, null],
				],
			},
		};
	}),
);

const outputPath = path.join(outputRoot, "formula_rewrite_codex_results.json");
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
