const addresses = ["Summary!D2", "Summary!D3", "Summary!D4", "Summary!D5", "Summary!F2", "Summary!G2", "Summary!H2"]

const cells = {}
for (const address of addresses) {
	const cell = await xlsx.readCell(wb, address)
	cells[address] = {
		value: cell.value,
		formula: cell.formula ?? null,
	}
}

return {
	tsv: await xlsx.readRangeTsv(wb, "Summary!D1:H5", {
		includeEmpty: true,
		includeFormulas: true,
	}),
	cells,
}
