# General code
* [x] Update `src/common/gas-utils/mapped-range.ts/resizeMappedRange` function to grow only on rows and spill over content.
* [ ] Add `readonly`, `Readonly<>`, `ReadonlyArray<>`, etc. **Everywhre**.
* [ ] Remove any direct use of `duplicateSheet:` and instead use `addNewSheet`.
* [ ] Remove any direct use of `buildAddNamedRangeRequest` or `addNamedRange:` and instead use `addNewNamedRange`.
* [ ] Replace string union types with string enums.
* [ ] Research markdown github "badges".
* [ ] Replace `sheet` with a stric version to avoid having to `?? 0` everywhere.

# Report initial creation
## Persistent Data
* [x] Number each student by groups
* [x] Generate sheet name from student names.
* [ ] Refactor to update ranges from within `buildTransferRequests` function.

## Assistance
* [x] Unbound `Validation` and `Conditional Format` from `faltasValidFormCond` cell.
* [x] Create constants for base row (4) and base column (9)
* [x] Protect Sheet!
* [x] Update formulas so it only takes 1/10
* [x] Set `BandedRange` for each subject, dynamic colors.
* [x] Remove format for rows between subjects.
* [x] Only allow write on unprotected period.
* [ ] Add vertical borders between the final periods.

## Student template
* [ ] Double check the first row is frozen (the student name) 
* [ ] Check editable ranges for comments

## Cards
* [ ] Add peekHeader (`newCardHeader`) to all Cards (`card.setPeekCardHeader`)

## Student Sheets

## Status



# General
* [x] Switch from using `formattedValue` to use `effectiveValue`.
* [x] Use `createRequiredGetter` and siplified names everywhere.
* [ ] Request "sheets.protectedRanges" to be able to edit them from the `protectedRangeId: number`;

# Menu 
## Reports
* [ ] Generate individual report
    - Dropdown to select student
    - Text field to set file name
* [ ] Generate all reports

## Periods
* [ ] Set current period
    - Radio select period (3 in total)
* [ ] Reset comments
* [ ] Protect/unprotect period
    - Checkbox select period(s) (3 in total)

## Calendar
* [ ] Add/remove day
    - Day (range) select
* [ ] Change period dates
    - Text field for each period date (3 in total)

## Students
* [ ] Add student
    - Full card with all the new studetn's data
* [ ] Remove student
    - Dropdown to select student
* [ ] Reorder student
    - Dropdown to select student
    - Dropdown? to select new position

## Grades
* [ ] Change general weights
    - Text field for each weight (3 in total)
* [ ] Change subject weights
    - Text field for each subject (about 15 in total)

## Subjects / Fields
* [ ] Rename Subject/Field
    - Dropdown to select subject/field
    - Text field to put new name
* [ ] Reorder Subject/Field
    - Dropdown to select subject/field
    - Dropdown? to select new position
* [ ] Add subject/field
    - Text field for name
    - Dropdown? for position
* [ ] Remove subject/field
    - Dropdown to select item


# Automatic Deploy on GitHub
* [ ] Update Template
* [ ] Create script / github action to auto-deploy
* [ ] Copy ~/.clasprc as a secret


# Function Flow

- 📄 [`buildDriveCard`](../src/drive-triggers.ts) - Main Drive entry point.
    - 📄 [`buildCreateSetupFileCard`](../src/setup/cards.ts) - Card Form to create a new Setup file 📋.
        - ⚡ [`onCreateSetupFile`](../src/setup/callbacks.ts) - Callback to create a new Setup file.
            - 🔀 [`createSetupFile`](../src/setup/crete-setup-file/index.ts) - Creates the new Setup file.
                - 🛠️ [`generateCalendar`](../src/setup/generate-calendar/index.ts) - Add the calendar to the Setup file.
    - 📄 TODO Make copy of setup
    - 📄 TODO Open in sheets to edit

- 📄 [`buildSheetsCard`](../src/sheet-triggers.ts) - Main Sheets entry point.
    - 📄 [`buildRequestAuthorizationCard`](../src/common/gas-parts/premade-cards.ts) - Ask the user for editing permission.
        - ⚡ [`onAskPermission`](../src/common/gas-parts/callbacks.ts) - Callback to show Google's default permission query.
    - 📄 [`buildEditSetupFileCard`](../src/setup/cards.ts) - Setup file 📋 editing and report creation card.
        - ⚡ [`onGenerateCalendar`](../src/setup/callbacks.ts) - Callback to regenerate the calendar.
            - 🛠️ [`generateCalendar`](../src/setup/generate-calendar/index.ts) - Add the calendar to the Setup file.
        - ⚡ [`onCopySetupFile`](../src/setup/callbacks.ts)
            - 🛠️ [`copySetupFile`](../src/setup/copy-setup-file/index.ts) - Creates a copy of the setup file with a given name.
        - ⚡ [`onInitializeReport`](../src/setup/callbacks.ts) - Uses the setup to initialize a report 📊
            - 🔀 [`initializeReport`](../src/setup/initialize-reports/index.ts) - Creates a new Report file 📊
                - 🛠️ `createReportFile` - Copy the template
                - 🛠️ [`fillPersistentData`](../src/setup/initialize-reports/persistent-data.ts) - Copy the Setup 📋 data over.
                - 🛠️ [`createAttendanceSheet`](../src/setup/initialize-reports/attendance.ts) - Create attendance sheet.
                - 🛠️ [`prepareStudentTemplate`](../src/setup/initialize-reports/student-template.ts) - Adapt the student template for the data of this group.
                - 🛠️ [`createStudentSheets`](../src/setup/initialize-reports/student-sheets.ts) - Create copies of the student template for each student.
                - 🛠️ [`prepareStatusSheet`](../src/setup/initialize-reports/status.ts) - Prepare the Status sheet to reach to each student sheet.
                - 🛠️ [`prepareSummarySheet`](../src/setup/initialize-reports/summary.ts) - Prepare the Summary to see the data of each student.

    - 📄 [`buildReportOptionsMainCard`](..src/report/cards.ts) - Main 📊 Report Admin Menu




📄 (page)
⚡ (zap)
⚒️ (hammer)
🔀 (twisted)
📋 (clipboard)
📊 (bar-chart)
📅 (date) 
🧑‍🎓 (student)
📚 (books)
⚖️ (balance-scale)