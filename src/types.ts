export type CellValue = string | number | boolean | undefined
export type Row = CellValue[]

export type Matcher = string | RegExp | ((name: string, idx: number, arr: string[]) => boolean)

export type Options = {
	handleBlankRow: 'throw' | 'excludeRow' | 'truncate'
	handleUncaughtCellError: 'throw' | 'excludeRow'
}

export type GenericWorkBookConfig<T extends GenericSheetsConfig> = {
	options?: Partial<Options>
	sheets: T
}

export type Converter<T> = (value: CellValue, meta: { reference: string; warnings: Warning[] }) => T

export type SchemaItem<T, E, B> = {
	match?: Matcher
	ifError?: E
	ifBlank?: B
	convert: Converter<T>
}

export type GenericSheetsConfig = Record<
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

export type Warning = {
	reference?: string
	message: string
	code?:
		| 'row_excluded_due_to_blanks'
		| 'rows_truncated_due_to_blanks'
		| 'row_excluded_due_to_cell_error'
		| 'cell_defaulted_due_to_error'
		| (string & { readonly customCode?: unique symbol })
}

export type Output<T extends GenericSheetsConfig> = {
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
