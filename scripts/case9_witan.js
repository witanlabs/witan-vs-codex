async function snapshot(label) {
	return {
		label,
		qty: await xlsx.readCell(wb, "Summary!B2"),
		price: await xlsx.readCell(wb, "Summary!B3"),
		spill: await xlsx.readRangeTsv(wb, "Summary!E2:J6", {
			includeEmpty: true,
			includeFormulas: true,
		}),
		data: await xlsx.readRangeTsv(wb, "Data!A1:D6", {
			includeEmpty: true,
			includeFormulas: true,
		}),
	}
}

const steps = []

steps.push(await snapshot("before"))

await xlsx.insertRowAfter(wb, "Data", 3, 1)
await xlsx.setCells(wb, [
	{ address: "Data!A4", value: "X" },
	{ address: "Data!B4", value: 10 },
	{ address: "Data!C4", value: 100 },
])
steps.push(await snapshot("after_insert_row"))

await xlsx.insertColumnAfter(wb, "Data", "B", 1)
await xlsx.setCells(wb, [
	{ address: "Data!C1", value: "Adj" },
	{ address: "Data!C2", value: 1000 },
	{ address: "Data!C3", value: 2000 },
	{ address: "Data!C4", value: 3000 },
	{ address: "Data!C5", value: 4000 },
	{ address: "Data!C6", value: 5000 },
])
steps.push(await snapshot("after_insert_col"))

await xlsx.deleteRows(wb, "Data", 5, 1)
steps.push(await snapshot("after_delete_row"))

await xlsx.deleteColumns(wb, "Data", "C", 1)
steps.push(await snapshot("after_delete_col"))

return { steps }
