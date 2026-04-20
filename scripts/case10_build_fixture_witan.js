const rows = [
	{ label: "TEXT multi-section accounting", expected: "($1,234.57)" },
	{ label: "TEXT conditional percent sections", expected: "37.5%" },
	{ label: "3D SUM", expected: 60 },
	{ label: "OFFSET SUM", expected: 10 },
	{ label: "INDIRECT SUM", expected: 10 },
	{ label: "MAP double SUM", expected: 20 },
	{ label: "REDUCE square SUM", expected: 30 },
	{ label: "LET qty*price SUM", expected: 300 },
	{ label: "XLOOKUP price for C", expected: 30 },
	{ label: "INDEX/XMATCH price for D", expected: 40 },
	{ label: "TEXTJOIN(MAP(...))", expected: "A:1,B:2,C:3,D:4" },
	{ label: "SUMPRODUCT filtered qty", expected: 9 },
	{ label: "SUM(SEQUENCE(3))", expected: 6 },
	{ label: "SUM(BYROW(...))", expected: 110 },
	{ label: "SUM(CHOOSECOLS(...))", expected: 10 },
	{ label: "SUM(TAKE(...,2))", expected: 3 },
	{ label: "SUM(DROP(...,2))", expected: 7 },
]

await xlsx.addSheet(wb, "Data")
await xlsx.addSheet(wb, "Jan")
await xlsx.addSheet(wb, "Feb")
await xlsx.addSheet(wb, "Mar")
await xlsx.addSheet(wb, "Config")
await xlsx.addSheet(wb, "Summary")

const cells = [
	{ address: "Data!A1", value: "Item" },
	{ address: "Data!B1", value: "Qty" },
	{ address: "Data!C1", value: "Price" },
	{ address: "Data!A2", value: "A" },
	{ address: "Data!B2", value: 1 },
	{ address: "Data!C2", value: 10 },
	{ address: "Data!A3", value: "B" },
	{ address: "Data!B3", value: 2 },
	{ address: "Data!C3", value: 20 },
	{ address: "Data!A4", value: "C" },
	{ address: "Data!B4", value: 3 },
	{ address: "Data!C4", value: 30 },
	{ address: "Data!A5", value: "D" },
	{ address: "Data!B5", value: 4 },
	{ address: "Data!C5", value: 40 },
	{ address: "Jan!A1", value: "Metric" },
	{ address: "Jan!B1", value: "Value" },
	{ address: "Jan!A2", value: "Sales" },
	{ address: "Jan!B2", value: 10 },
	{ address: "Feb!A1", value: "Metric" },
	{ address: "Feb!B1", value: "Value" },
	{ address: "Feb!A2", value: "Sales" },
	{ address: "Feb!B2", value: 20 },
	{ address: "Mar!A1", value: "Metric" },
	{ address: "Mar!B1", value: "Value" },
	{ address: "Mar!A2", value: "Sales" },
	{ address: "Mar!B2", value: 30 },
	{ address: "Config!A1", value: "IndirectSheet" },
	{ address: "Config!B1", value: "Data" },
	{ address: "Config!A2", value: "IndirectRange" },
	{ address: "Config!B2", value: "B2:B5" },
	{ address: "Summary!A1", value: "Formula" },
	{ address: "Summary!B1", value: "Expected" },
	{ address: "Summary!C1", value: "Actual" },
]

rows.forEach((row, index) => {
	const sheetRow = index + 2
	cells.push(
		{ address: `Summary!A${sheetRow}`, value: row.label },
		{ address: `Summary!B${sheetRow}`, value: row.expected },
	)
})

await xlsx.setCells(wb, cells)

return {
	sheets: await xlsx.listSheets(wb),
	summary: await xlsx.readRangeTsv(wb, "Summary!A1:C18", {
		includeEmpty: true,
		includeFormulas: true,
	}),
}
