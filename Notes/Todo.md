# General code
* [x] Update `src/common/gas-utils/mapped-range.ts/resizeMappedRange` function to grow only on rows and spill over content.
* [ ] Add `readonly`, `Readonly<>`, `ReadonlyArray<>`, etc. **Everywhre**.
* [ ] Remove any direct use of `duplicateSheet:` and instead use `addNewSheet`.
* [ ] Remove any direct use of `buildAddNamedRangeRequest` or `addNamedRange:` and instead use `addNewNamedRange`.
* [ ] Replace string union types with string enums.
* [ ] Research markdown github "badges".

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

## Student Sheets



# General
* [x] Switch from using `formattedValue` to use `effectiveValue`.
* [x] Use `createRequiredGetter` and siplified names everywhere.
* [ ] Request "sheets.protectedRanges" to be able to edit them from the `protectedRangeId: number`;

# Menu
## Assistance
* [ ] Recalculate periods

## Persistant Data
* [ ] Normalize weights (for grades and subjects) 
* [ ] Change weights


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

    - 📄 TODO Report file editing (this is the biggest UI entry point)



📄 (page)
⚡ (zap)
⚒️ (hammer)
🔀 (twisted)
📋 (clipboard)
📊 (bar-chart)