# Sheet to Schema

Convert an XLSX spreadsheet into a TypeScript-aware object using a schema:

```ts
import { sheetToSchema } from 'sheet-to-schema/mod.ts'
import { xlsxDateTime } from 'sheet-to-schema/dates.ts'

const output = sheetToSchema(await Deno.readFile('/path/to/file.xlsx'), {
    sheets: {
        main: {
            match: /sheet.?1/i,
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

// for illustration purposes - types are automatically inferred based on schema
type Data = {
    firstName: string
    lastName: string
    dob: Temporal.PlainDateTime | null
}[]

const data: Data = output.results.main
```

Uses SheetJS as a dependency. `dates.ts` additionally uses `Temporal` polyfill as a dependency.
