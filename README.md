# Time Guardian

Time Guardian is an attendance compiler for combining Fingerprint and Online Excel exports into one HR-ready attendance workbook.

It was built for the October 2025 attendance flow, but the parsing and calculation logic are reusable for similar monthly reports.

## What it does

- Upload a Fingerprint Excel file and an Online Excel file
- Match employees by name, including first-name and last-name overlap
- Merge the earliest clock-in and latest clock-out from both sources
- Apply break, flexi, tardiness, and overtime rules
- Generate a compiled Excel file with one worksheet per employee

## Attendance rules

- Break deduction
  - Monday to Thursday: 12:00 - 12:30
  - Friday: 11:30 - 13:00
- Flexi time
  - Flexi 1: 08:00 - 08:15
  - Flexi 2: 08:15 - 08:30
  - After 08:30: tardiness
- Overtime
  - Monday to Thursday: starts at 17:30
  - Friday: starts at 18:00

## Tech stack

- Vite
- React
- TypeScript
- shadcn/ui
- Tailwind CSS
- SheetJS (`xlsx`)

## Local setup

```sh
npm install
npm run dev
```

## Scripts

- `npm run dev` - start the app locally
- `npm run build` - build for production
- `npm run lint` - run ESLint
- `npm run test` - run the test suite

## Notes

- The workbook output is designed to match the manual compiled attendance format as closely as possible.
- The report period is inferred from the uploaded files.
- If the two source files contain different months, the current flow is not guaranteed to merge them correctly.

