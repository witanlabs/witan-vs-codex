async function snapshot(label) {
	return {
		label,
		sheets: await xlsx.listSheets(wb),
		qty: await xlsx.readCell(wb, "Summary!B2"),
		total: await xlsx.readCell(wb, "Summary!B3"),
		spill: await xlsx.readRangeTsv(wb, "Summary!D2:F4", {
			includeEmpty: true,
			includeFormulas: true,
		}),
	}
}

const before = await snapshot("before")
await xlsx.renameSheet(wb, "Data", "Renamed")
const after = await snapshot("after")

return { before, after }
