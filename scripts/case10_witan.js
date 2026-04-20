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
]

async function snapshot(label) {
	const rows = []
	for (const spec of formulaSpecs) {
		const actual = await xlsx.readCell(wb, `Summary!C${spec.row}`)
		const expected = await xlsx.readCell(wb, `Summary!B${spec.row}`)
		const itemLabel = await xlsx.readCell(wb, `Summary!A${spec.row}`)
		rows.push({
			key: spec.key,
			label: itemLabel.value,
			expected: expected.value,
			formula: actual.formula ?? null,
			value: actual.value,
		})
	}
	return {
		label,
		rows,
		table: await xlsx.readRangeTsv(wb, "Summary!A1:C18", {
			includeEmpty: true,
			includeFormulas: true,
		}),
	}
}

const before = await snapshot("before")

await xlsx.setCells(
	wb,
	formulaSpecs.map((spec) => ({
		address: `Summary!C${spec.row}`,
		formula: spec.formula,
	})),
)

const after = await snapshot("after")

return { before, after }
