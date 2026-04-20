const sheets = await xlsx.listSheets(wb)
const firstSheet = sheets[0]?.sheet ?? sheets[0]?.name
const cell = await xlsx.readCell(wb, { sheet: firstSheet, row: 3, col: 2 })

return {
	sheets,
	firstSheet,
	cell,
}
