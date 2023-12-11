import * as XLSX from 'https://raw.githubusercontent.com/clearlylocal/xlsx/9910080/xlsx.mjs'
import { quoteSheetName, toExcelCol } from './references.ts'
import type {
	CellValue,
	Converter,
	GenericSheetsConfig,
	GenericWorkBookConfig,
	Options,
	Output,
	Row,
	Warning,
} from './types.ts'
export type { CellValue, GenericSheetsConfig, GenericWorkBookConfig }

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

const defaultOptions: Options = {
	handleBlankRow: 'throw',
	handleUncaughtCellError: 'throw',
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

			const headingInfo = Object.entries(schema).map(([key, val]) => ({ ...val, key }))
			type HeadingInfo = typeof headingInfo[number]
			// const headingMatchers = headingInfo.map((v) => v.match)

			let headerRowIndex = -1
			let foundLength = 0
			let foundHeaders: string[] = []
			for (const [idx, row] of rows.slice(minIndex, maxIndex).entries()) {
				const found = headingInfo.filter((m) => row.map(String).find(m.match))

				if (found.length > foundLength) {
					foundLength = found.length
					foundHeaders = found.map((x) => x.key)
				}

				if (found.length === headingInfo.length) {
					headerRowIndex = idx
					break
				}
			}

			if (headerRowIndex === -1) {
				throw new TypeError(
					`Headers ${
						new Intl.ListFormat('en-US').format(
							Object.keys(schema).filter((k) => !foundHeaders.includes(k)).map((x) => JSON.stringify(x)),
						)
					} missing for sheet ${JSON.stringify(k)}`,
				)
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

const DEFAULT_READ_OPTS = { dense: true }

export function sheetToSchema<T extends GenericSheetsConfig>(config: GenericWorkBookConfig<T>): (
	xlsxBin: Uint8Array | ArrayBuffer,
) => Expand<Output<T>>
export function sheetToSchema<T extends GenericSheetsConfig>(
	xlsxBin: Uint8Array | ArrayBuffer,
	config: GenericWorkBookConfig<T>,
): Expand<Output<T>>
// deno-lint-ignore no-explicit-any
export function sheetToSchema(...args: any[]) {
	switch (args.length) {
		case 1: {
			const [config] = args as [GenericWorkBookConfig<GenericSheetsConfig>]
			return (xlsxBin: Uint8Array | ArrayBuffer) => {
				const wb = XLSX.read(xlsxBin, { ...DEFAULT_READ_OPTS, ...config.readOptions })
				return getWorkbookData(wb, config)
			}
		}
		case 2: {
			const [xlsxBin, config] = args as [Uint8Array | ArrayBuffer, GenericWorkBookConfig<GenericSheetsConfig>]
			const wb = XLSX.read(xlsxBin, { ...DEFAULT_READ_OPTS, ...config.readOptions })
			return getWorkbookData(wb, config)
		}
		default: {
			throw new RangeError(`Wrong number of arguments: ${args.length}`)
		}
	}
}

type Expand<T> = T extends Record<string, unknown> ? T extends infer O ? { [K in keyof O]: Expand<O[K]> } : never
	: T
