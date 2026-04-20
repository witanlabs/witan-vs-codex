import fs from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";

import { FileBlob, SpreadsheetFile } from "@oai/artifact-tool";

const scriptDir = path.dirname(fileURLToPath(import.meta.url));
const repoRoot = path.resolve(scriptDir, "..");
const fixturesRoot = path.join(repoRoot, "fixtures");
const outputRoot = path.join(repoRoot, "outputs");
const selected = new Set(process.argv.slice(2));

await fs.mkdir(outputRoot, { recursive: true });

function shouldRun(name) {
  return selected.size === 0 || selected.has(name);
}

function cellValue(ws, address) {
  return ws.getRange(address).values?.[0]?.[0] ?? null;
}

function cellFormula(ws, address) {
  return ws.getRange(address).formulas?.[0]?.[0] ?? null;
}

function rangeValues(ws, address) {
  return ws.getRange(address).values ?? [];
}

function rangeFormulas(ws, address) {
  return ws.getRange(address).formulas ?? [];
}

function simplifyThreads(proto) {
  const people = new Map((proto.people ?? []).map((person) => [person.id, person]));
  return (proto.threads ?? []).map((thread) => {
    const target =
      thread.target?.spreadsheetCell ??
      thread.target?.cell ??
      thread.target?.range ??
      {};
    return {
      id: thread.id ?? null,
      sheet: target.sheetName ?? null,
      address: target.address ?? target.startAddress ?? null,
      status: thread.status ?? null,
      resolvedAt: thread.resolvedAt ?? null,
      resolvedBy: people.get(thread.resolvedBy ?? "")?.displayName ?? thread.resolvedBy ?? null,
      comments: (thread.comments ?? []).map((comment) => ({
        authorId: comment.authorId ?? null,
        author: people.get(comment.authorId ?? "")?.displayName ?? null,
        text: comment.body?.plainText ?? comment.text ?? null,
        createdAt: comment.createdAt ?? null,
      })),
    };
  });
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
  if (!shouldRun(name)) {
    return null;
  }
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
  await runCase("case2", async () => {
    const workbook = await importWorkbook("circular.xlsx");
    const inputs = workbook.worksheets.getItem("Inputs");
    const model = workbook.worksheets.getItem("Model");

    inputs.getRange("B4").values = [[0.2]];
    workbook.recalculate();

    const outputPath = await exportWorkbook(workbook, "case2_codex.xlsx");
    return {
      outputPath,
      fixture: "fixtures/circular.xlsx",
      artifact: {
        bonusRate: cellValue(inputs, "B4"),
        profit: cellValue(model, "B3"),
        bonus: cellValue(model, "B4"),
        netIncome: cellValue(model, "B7"),
      },
      expected: {
        bonusRate: 0.2,
        profit: 50000,
        bonus: 10000,
        netIncome: 35000,
      },
    };
  }),
);

results.push(
  await runCase("case3", async () => {
    const input = await FileBlob.load(path.join(fixturesRoot, "formulas.xls"));
    const workbook = await SpreadsheetFile.importXlsx(input);
    const firstSheet = workbook.worksheets.getItemAt(0);
    const outputPath = await exportWorkbook(workbook, "case3_codex.xlsx");
    return {
      outputPath,
      fixture: "fixtures/formulas.xls",
      artifact: {
        sheetName: firstSheet.name,
        b3: cellValue(firstSheet, "B3"),
      },
      expected: {
        b3: 4,
      },
    };
  }),
);

results.push(
  await runCase("case4", async () => {
    const workbook = await importWorkbook("review.xlsx");
    const data = workbook.worksheets.getItem("Data");
    const before = simplifyThreads(workbook.comments.toProto());

    const thread = workbook.comments.addThread(
      { cell: data.getRange("B3") },
      "Verified against ledger",
      { author: { displayName: "Auditor", initials: "AU" } },
    );
    thread.resolve({ displayName: "Auditor", initials: "AU" });

    const outputPath = await exportWorkbook(workbook, "case4_codex.xlsx");
    return {
      outputPath,
      fixture: "fixtures/review.xlsx",
      artifact: {
        before,
        after: simplifyThreads(workbook.comments.toProto()),
      },
      expected: {
        addresses: ["B2", "C2", "B3"],
        threadCount: 3,
      },
    };
  }),
);

results.push(
  await runCase("case5", async () => {
    const workbook = await importWorkbook("report.xlsx");
    const summary = workbook.worksheets.getItem("Summary");

    summary.getRange("D2").formulas = [[
      "=UNIQUE(FILTER(Raw!A2:A13, Raw!B2:B13>0))",
    ]];
    workbook.recalculate();

    const outputPath = await exportWorkbook(workbook, "case5_codex.xlsx");
    return {
      outputPath,
      fixture: "fixtures/report.xlsx",
      artifact: {
        values: rangeValues(summary, "D2:D10").flat(),
        formulas: rangeFormulas(summary, "D2:D10").flat(),
      },
      expected: {
        values: ["Food", "Rent", "Supplies", "Travel"],
      },
    };
  }),
);

results.push(
  await runCase("case7", async () => {
    const workbook = await importWorkbook("report.xlsx");
    const summary = workbook.worksheets.getItem("Summary");

    summary.getRange("D1:H5").values = [
      ["Unique pos cats", null, "Count", "Food matches", "Joined"],
      [null, null, null, null, null],
      [null, null, null, null, null],
      [null, null, null, null, null],
      [null, null, null, null, null],
    ];
    summary.getRange("D2").formulas = [[
      "=UNIQUE(FILTER(Raw!A2:A13, Raw!B2:B13>0))",
    ]];
    summary.getRange("F2").formulas = [["=COUNTA(Summary!D2#)"]];
    summary.getRange("G2").formulas = [['=COUNTIF(Summary!D2#, "Food")']];
    summary.getRange("H2").formulas = [['=TEXTJOIN(", ", TRUE, Summary!D2#)']];
    workbook.recalculate();

    const outputPath = await exportWorkbook(workbook, "case7_codex.xlsx");
    return {
      outputPath,
      fixture: "fixtures/report.xlsx",
      artifact: {
        values: rangeValues(summary, "D1:H5"),
        formulas: rangeFormulas(summary, "D1:H5"),
      },
      expected: {
        spillValues: ["Food", "Rent", "Supplies", "Travel"],
        count: 4,
        foodMatches: 1,
        joined: "Food, Rent, Supplies, Travel",
      },
    };
  }),
);

const filteredResults = results.filter(Boolean);
const summaryPath = path.join(outputRoot, "artifact_results.json");

await fs.writeFile(
  summaryPath,
  `${JSON.stringify(
    {
      generatedAt: new Date().toISOString(),
      repoRoot,
      outputRoot,
      results: filteredResults,
    },
    null,
    2,
  )}\n`,
);

console.log(JSON.stringify({ summaryPath, results: filteredResults }, null, 2));
