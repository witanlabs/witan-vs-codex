await xlsx.addSheet(wb, "Data")
await xlsx.addSheet(wb, "Summary")

await xlsx.setCells(wb, [
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
	{ address: "Summary!B1", value: "Qty sum" },
	{ address: "Summary!B2", formula: "=SUM(Data!B2:B5)" },
	{ address: "Summary!B3", formula: "=SUM(Data!C2:C5)" },
	{ address: "Summary!E1", value: "Transposed Data" },
	{ address: "Summary!E2", formula: "=TRANSPOSE(Data!A2:C5)" },
])

return {
	sheets: await xlsx.listSheets(wb),
	summary: await xlsx.readRangeTsv(wb, "Summary!B1:J5", {
		includeEmpty: true,
		includeFormulas: true,
	}),
}
