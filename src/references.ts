// JS port from https://stackoverflow.com/a/48984697
const divmodExcel = (n: number) => {
	const a = Math.floor(n / 26)
	const b = n % 26

	return b === 0 ? [a - 1, b + 26] : [a, b]
}

const uppercaseAlpha = Array.from({ length: 26 }, (_, i) => String.fromCodePoint(i + 'A'.codePointAt(0)!))

export const toExcelCol = (n: number) => {
	const chars = []

	let d: number
	while (n > 0) {
		;[n, d] = divmodExcel(n)
		chars.unshift(uppercaseAlpha[d - 1])
	}
	return chars.join('')
}

export function quoteSheetName(sheetName: string) {
	return `'${sheetName.replaceAll('\'', '\'\'')}'`
}
