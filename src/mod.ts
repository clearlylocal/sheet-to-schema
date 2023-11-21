import * as XLSX from 'https://raw.githubusercontent.com/clearlylocal/xlsx/9910080/xlsx.mjs'
import { quoteSheetName, toExcelCol } from './references.ts'
import type { CellValue } from './types.ts'

type Matcher = string | RegExp | ((name: string, idx: number, arr: string[]) => boolean)
type Row = CellValue[]

function isBlank(val: CellValue) {
	return val == null || val === '' || val === 0
}

function isContentful(val: CellValue) {
	return !isBlank(val)
}

function sheetTo2dArray(sheet: XLSX.Sheet) {
	const rows = XLSX.utils.sheet_to_json<Row>(sheet, { header: 1 })
	const maxRow = rows.findLastIndex((row) => row.some(isContentful))

	return rows.slice(0, maxRow + 1)
}

type Options = {
	handleBlankRow: 'throw' | 'excludeRow' | 'truncate'
	handleUncaughtCellError: 'throw' | 'excludeRow'
}

type GenericWorkBookConfig<T extends GenericSheetsConfig> = {
	options?: Partial<Options>
	sheets: T
}

const defaultOptions: Options = {
	handleBlankRow: 'throw',
	handleUncaughtCellError: 'throw',
}

type Converter<T> = (value: CellValue, meta: { reference: string; warnings: Warning[] }) => T

type SchemaItem<T, E, B> = {
	match?: Matcher
	ifError?: E
	ifBlank?: B
	convert: Converter<T>
}

type GenericSheetsConfig = Record<
	string,
	{
		match?: Matcher
		headerRow?: {
			minIndex: number
			maxIndex: number
		}
		schema: Record<string, SchemaItem<unknown, unknown, unknown>>
	}
>

type Warning = {
	reference?: string
	message: string
	code?:
		| 'row_excluded_due_to_blanks'
		| 'rows_truncated_due_to_blanks'
		| 'row_excluded_due_to_cell_error'
		| 'cell_defaulted_due_to_error'
		| (string & { readonly customCode?: unique symbol })
}

type Output<T extends GenericSheetsConfig> = {
	results: {
		[SheetKey in keyof T]: {
			[SchemaKey in keyof T[SheetKey]['schema']]:
				| ReturnType<T[SheetKey]['schema'][SchemaKey]['convert']>
				| (
					T[SheetKey]['schema'][SchemaKey] extends { ifError: unknown }
						? T[SheetKey]['schema'][SchemaKey]['ifError']
						: never
				)
				| (
					T[SheetKey]['schema'][SchemaKey] extends { ifBlank: unknown }
						? T[SheetKey]['schema'][SchemaKey]['ifBlank']
						: never
				)
		}[]
	}
	warnings: Warning[]
}

const MIN_HEADER_ROW_INDEX = 0
const MAX_HEADER_ROW_INDEX = 100

function getWorkbookData<T extends GenericSheetsConfig>(
	wb: XLSX.WorkBook,
	params: GenericWorkBookConfig<T>,
): Expand<Output<T>> {
	const warnings: Warning[] = []

	const options: Options = {
		...defaultOptions,
		...params.options,
	}

	const results = Object.fromEntries(
		Object.entries(params.sheets).map(([k, sheetConfig]) => {
			const match = sheetConfig.match ?? k
			const sheetMatcher = typeof match === 'string'
				? (x: string) => x === match
				: match instanceof RegExp
				? match.test.bind(match)
				: match

			const sheetName = wb.SheetNames.find(sheetMatcher)

			if (!sheetName) throw new TypeError(`No matching sheet found for ${JSON.stringify(k)}`)

			const rows = sheetTo2dArray(wb.Sheets[sheetName]!)

			const { minIndex, maxIndex } = sheetConfig.headerRow ?? {
				minIndex: MIN_HEADER_ROW_INDEX,
				maxIndex: MAX_HEADER_ROW_INDEX,
			}

			const schema = Object.fromEntries(
				Object.entries(sheetConfig.schema).map(([k, v]) => {
					const match = v.match ?? k
					const headingMatcher = typeof match === 'string'
						? (x: string) => x === match
						: match instanceof RegExp
						? match.test.bind(match)
						: match

					const convert: Converter<unknown> = (
						value: CellValue,
						meta: { reference: string; warnings: Warning[] },
					) => {
						const { reference, warnings } = meta
						try {
							if (isBlank(value)) {
								if (Object.hasOwn(v, 'ifBlank')) {
									return v.ifBlank
								}

								switch (v.convert) {
									case String:
										return ''
									case Boolean:
										return false
									case Number:
										return 0
									case BigInt:
										return 0n
									default:
										return v.convert(value, meta)
								}
							}

							return (Object.hasOwn(v, 'ifBlank')) && isBlank(value) ? v.ifBlank : v.convert(value, meta)
						} catch (e) {
							if (Object.hasOwn(v, 'ifError')) {
								warnings.push({
									reference,
									code: 'cell_defaulted_due_to_error',
									message: e?.message ?? String(e),
								})

								return v.ifError
							} else {
								throw e
							}
						}
					}

					return [k, {
						match: headingMatcher,
						convert,
					}] as const
				}),
			)

			const headingMatchers = Object.values(schema).map((v) => v.match)

			const headerRowIndex = rows.slice(minIndex, maxIndex).findIndex((row) => {
				return headingMatchers.every((m) => row.map(String).some(m))
			})

			if (headerRowIndex === -1) {
				throw new TypeError(`No matching header row for sheet ${k}`)
			}

			const _headings = rows[headerRowIndex].map(String)
			const headings = Object.fromEntries(
				Object.entries(schema).map(([k, v]) => {
					return [k, _headings.findIndex((...args) => v.match(...args))]
				}),
			)

			const items: Record<string, unknown>[] = []

			eachRow: for (const [rowIndex, row] of rows.slice(headerRowIndex + 1).entries()) {
				const excelRowNumber = rowIndex + headerRowIndex + 2

				if (row.every(isBlank)) {
					switch (options.handleBlankRow) {
						case 'excludeRow': {
							warnings.push({
								reference: `${quoteSheetName(sheetName)}!${excelRowNumber}:${excelRowNumber}`,
								code: 'row_excluded_due_to_blanks',
								message: `Row ${excelRowNumber} is blank and was excluded from results`,
							})

							continue eachRow
						}
						case 'truncate': {
							warnings.push({
								reference: `${quoteSheetName(sheetName)}!${excelRowNumber}:${excelRowNumber}`,
								code: 'rows_truncated_due_to_blanks',
								message:
									`Row ${excelRowNumber} is blank, and results were truncated started from this row`,
							})

							break eachRow
						}
						case 'throw':
						default: {
							throw new TypeError(`Row ${excelRowNumber} is blank`)
						}
					}
				}

				items.push(Object.fromEntries(
					Object.entries(headings).map(([k, v]) => {
						return [
							k,
							schema[k].convert(row[v], {
								reference: `${quoteSheetName(sheetName)}!${toExcelCol(v + 1)}${excelRowNumber}`,
								warnings,
							}),
						]
					}),
				))
			}

			return [k, items]
		}),
	)

	return { results: results as Expand<Output<T>['results']>, warnings }
}

export function sheetToSchema<T extends GenericSheetsConfig>(
	xlsxBin: Uint8Array | ArrayBuffer,
	config: GenericWorkBookConfig<T>,
) {
	const wb = XLSX.read(xlsxBin)

	return getWorkbookData(wb, config)
}

type Expand<T> = T extends Record<string, unknown> ? T extends infer O ? { [K in keyof O]: Expand<O[K]> } : never
	: T
