import { sheetToSchema } from '../src/mod.ts'
import { xlsxDateTime } from '../src/dates.ts'
import 'https://esm.sh/v131/temporal-polyfill@0.1.1/dist/global.mjs'
import * as z from 'https://esm.sh/v131/zod@3.21.4'
import { assertEquals, assertExists } from 'https://deno.land/std@0.207.0/assert/mod.ts'

class EmailAddress extends String {}

// for debugging purposes
for (
	const Class of [
		Temporal.Instant,
		Temporal.Calendar,
		Temporal.PlainDate,
		Temporal.PlainDateTime,
		Temporal.Duration,
		Temporal.PlainMonthDay,
		Temporal.PlainTime,
		Temporal.TimeZone,
		Temporal.PlainYearMonth,
		Temporal.ZonedDateTime,

		Intl.Locale,
	]
) {
	// deno-lint-ignore no-explicit-any
	;(Class.prototype as any)[Symbol.for('Deno.customInspect')] ??= function (this: any) {
		return `\x1b[36m${this[Symbol.toStringTag] ?? this.constructor?.name} <${this.toString()}>\x1b[0m`
	}
}

Deno.test('tasks-test-data', async (t) => {
	const output = sheetToSchema(await Deno.readFile('./tests/data/tasks-test-data.xlsx'), {
		options: {
			handleBlankRow: 'excludeRow',
		},
		sheets: {
			main: {
				match: /example/i,
				schema: {
					no: {
						match: /no\./i,
						convert: Number,
					},
					writer: {
						match: /writer/i,
						convert(x) {
							if (typeof x !== 'string' || !x.includes('@')) {
								throw new TypeError(`Expected ${JSON.stringify(x)} to be an email address`)
							}

							return new EmailAddress(new URL(`mailto:${x}`).pathname)
						},
						ifError: null,
					},
					category: {
						match: /category/i,
						convert: String,
					},
					keyword: {
						match: /keyword/i,
						convert: String,
					},
					url: {
						match: /url/i,
						convert(x) {
							return new URL(x as string)
						},
						ifError: null,
					},
					locale: {
						match: /locale/i,
						convert(x) {
							return new Intl.Locale(x as string)
						},
						ifError: null,
					},
					wordCountTarget: {
						match: /word.?count.?target/i,
						convert(x, { reference, warnings }) {
							if (x == null) return 400

							switch (x) {
								case 400:
								case 800:
									return x
								default:
									warnings.push({
										reference,
										code: 'CUSTOM_CODE',
										message: `Expected 400 or 800. Got ${x}`,
									})
									return 400
									// throw new TypeError(`Must be 400 or 800`)
							}
						},
					},
					due: {
						match: /due/i,
						convert: xlsxDateTime,
						ifError: null,
					},
				},
			},
			other: {
				match: /other/i,
				schema: {
					firstName: {
						match: /first.?name/i,
						convert: String,
					},
					lastName: {
						match: /last.?name/i,
						convert: String,
					},
					dob: {
						match: /date.?of.?birth/i,
						convert: xlsxDateTime,
						ifError: null,
					},
				},
			},
		},
	})

	const Results = z.object({
		main: z.object({
			no: z.number().or(z.nan()),
			writer: z.instanceof(EmailAddress).nullable(),
			category: z.string(),
			keyword: z.string(),
			url: z.instanceof(URL).nullable(),
			locale: z.instanceof(Intl.Locale).nullable(),
			wordCountTarget: z.literal(400).or(z.literal(800)),
			due: z.instanceof(Temporal.PlainDateTime).nullable(),
		}).array(),
		other: z.object({
			firstName: z.string(),
			lastName: z.string(),
			dob: z.instanceof(Temporal.PlainDateTime).nullable(),
		}).array(),
	})

	await t.step('results', async (t) => {
		const { results } = output

		await t.step('runtime types', () => {
			Results.parse(results)
		})

		await t.step('compile-time types (bidirectional)', () => {
			// bidirectional compile-time type check
			const _x: z.infer<typeof Results> = results
			const _y: typeof results = _x
		})
	})

	await t.step('warnings', async (t) => {
		const { warnings } = output

		await t.step('codes and references', () => {
			for (
				const [reference, code] of [
					['\'example\'!M11', 'cell_defaulted_due_to_error'],
					['\'example\'!M12', 'cell_defaulted_due_to_error'],

					['\'example\'!33:33', 'row_excluded_due_to_blanks'],

					['\'example\'!B32', 'cell_defaulted_due_to_error'],

					['\'example\'!G22', 'cell_defaulted_due_to_error'],
					['\'example\'!G23', 'cell_defaulted_due_to_error'],
					['\'example\'!G24', 'cell_defaulted_due_to_error'],
					['\'example\'!G25', 'cell_defaulted_due_to_error'],

					['\'example\'!N15', 'CUSTOM_CODE'],

					['\'example\'!O6', 'cell_defaulted_due_to_error'],
					['\'example\'!O21', 'cell_defaulted_due_to_error'],
				]
			) {
				assertExists(warnings.find((w) => w.reference === reference && w.code === code))
			}
		})
	})

	await t.step('snapshot', async () => {
		const serialized = JSON.stringify(output, (_, v) => {
			return v instanceof Intl.Locale ? v.toString() : v
		}, '\t')

		// // to update snapshot:
		// await Deno.writeTextFile('./tests/data/tasks-test-snapshot.json', serialized)

		const serializable = JSON.parse(serialized)

		const snapshot = JSON.parse(await Deno.readTextFile('./tests/data/tasks-test-snapshot.json'))

		assertEquals(serializable, snapshot)
	})
})

Deno.test('tasks-test-data curried', async (t) => {
	const getData = sheetToSchema({
		options: {
			handleBlankRow: 'excludeRow',
		},
		sheets: {
			main: {
				match: /example/i,
				schema: {
					no: {
						match: /no\./i,
						convert: Number,
					},
					writer: {
						match: /writer/i,
						convert(x) {
							if (typeof x !== 'string' || !x.includes('@')) {
								throw new TypeError(`Expected ${JSON.stringify(x)} to be an email address`)
							}

							return new EmailAddress(new URL(`mailto:${x}`).pathname)
						},
						ifError: null,
					},
					category: {
						match: /category/i,
						convert: String,
					},
					keyword: {
						match: /keyword/i,
						convert: String,
					},
					url: {
						match: /url/i,
						convert(x) {
							return new URL(x as string)
						},
						ifError: null,
					},
					locale: {
						match: /locale/i,
						convert(x) {
							return new Intl.Locale(x as string)
						},
						ifError: null,
					},
					wordCountTarget: {
						match: /word.?count.?target/i,
						convert(x, { reference, warnings }) {
							if (x == null) return 400

							switch (x) {
								case 400:
								case 800:
									return x
								default:
									warnings.push({
										reference,
										code: 'CUSTOM_CODE',
										message: `Expected 400 or 800. Got ${x}`,
									})
									return 400
									// throw new TypeError(`Must be 400 or 800`)
							}
						},
					},
					due: {
						match: /due/i,
						convert: xlsxDateTime,
						ifError: null,
					},
				},
			},
			other: {
				match: /other/i,
				schema: {
					firstName: {
						match: /first.?name/i,
						convert: String,
					},
					lastName: {
						match: /last.?name/i,
						convert: String,
					},
					dob: {
						match: /date.?of.?birth/i,
						convert: xlsxDateTime,
						ifError: null,
					},
				},
			},
		},
	})

	const output = getData(await Deno.readFile('./tests/data/tasks-test-data.xlsx'))

	const Results = z.object({
		main: z.object({
			no: z.number().or(z.nan()),
			writer: z.instanceof(EmailAddress).nullable(),
			category: z.string(),
			keyword: z.string(),
			url: z.instanceof(URL).nullable(),
			locale: z.instanceof(Intl.Locale).nullable(),
			wordCountTarget: z.literal(400).or(z.literal(800)),
			due: z.instanceof(Temporal.PlainDateTime).nullable(),
		}).array(),
		other: z.object({
			firstName: z.string(),
			lastName: z.string(),
			dob: z.instanceof(Temporal.PlainDateTime).nullable(),
		}).array(),
	})

	await t.step('results', async (t) => {
		const { results } = output

		await t.step('runtime types', () => {
			Results.parse(results)
		})

		await t.step('compile-time types (bidirectional)', () => {
			// bidirectional compile-time type check
			const _x: z.infer<typeof Results> = results
			const _y: typeof results = _x
		})
	})

	await t.step('warnings', async (t) => {
		const { warnings } = output

		await t.step('codes and references', () => {
			for (
				const [reference, code] of [
					['\'example\'!M11', 'cell_defaulted_due_to_error'],
					['\'example\'!M12', 'cell_defaulted_due_to_error'],

					['\'example\'!33:33', 'row_excluded_due_to_blanks'],

					['\'example\'!B32', 'cell_defaulted_due_to_error'],

					['\'example\'!G22', 'cell_defaulted_due_to_error'],
					['\'example\'!G23', 'cell_defaulted_due_to_error'],
					['\'example\'!G24', 'cell_defaulted_due_to_error'],
					['\'example\'!G25', 'cell_defaulted_due_to_error'],

					['\'example\'!N15', 'CUSTOM_CODE'],

					['\'example\'!O6', 'cell_defaulted_due_to_error'],
					['\'example\'!O21', 'cell_defaulted_due_to_error'],
				]
			) {
				assertExists(warnings.find((w) => w.reference === reference && w.code === code))
			}
		})
	})
})
