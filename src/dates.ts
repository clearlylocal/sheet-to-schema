import { Temporal } from 'https://esm.sh/v131/@js-temporal/polyfill@0.4.4'
import type { CellValue } from './types.ts'

/**
 * @param value An Excel value, stored as a number, and possibly formatted as a date and/or time when viewed
 * within Excel. Represents the number of days since January 1, 1900, with any fractional part indicating time of day.
 *
 * Note: ALL datetimes in Excel are technically `PlainDateTime`s as they carry no time zone info (!!) so we will also
 * need to provide the user's `TimeZone` object or string if we later want to convert to a `ZonedDateTime` using
 * `PlainDateTime#toZonedDateTime`.
 *
 * @throws {TypeError} at runtime if `value` is not a number. We allow non-number cell values for convenience when used
 * alongside `sheetToSchema`.
 *
 * @example
 *
 * const plainDateTime = xlsxDateTime(45247.5)
 * 	=> PlainDateTime <'2023-11-17T12:00:00'>
 * plainDateTime.toZonedDateTime('Asia/Shanghai')
 * 	=> ZonedDateTime <'2023-11-17T12:00:00+08:00[Asia/Shanghai]'>
 */
export function xlsxDateTime(value: CellValue) {
	if (typeof value !== 'number') {
		throw new TypeError(`Cannot convert ${JSON.stringify(value)} to a PlainDateTime: expected number`)
	}

	// See See https://stackoverflow.com/questions/6154953/excel-date-to-unix-timestamp for conversion logic
	return new Temporal.Instant(BigInt(Math.round((value - 25569) * 86400 * 1000)) * 10n ** 6n)
		.toZonedDateTimeISO('UTC')
		.toPlainDateTime()
}
