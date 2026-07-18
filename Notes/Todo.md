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

## Student Sheets

## Status



# General
* [x] Switch from using `formattedValue` to use `effectiveValue`.
* [x] Use `createRequiredGetter` and siplified names everywhere.
* [ ] Request "sheets.protectedRanges" to be able to edit them from the `protectedRangeId: number`;

# Menu
## Assistance
* [ ] Recalculate periods
* [ ] Insert / Delete days
* [ ] Add / Remove / Reorder students

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
                - 🛠️ [`prepareStudentTemplate`](../src/setup/initialize-reports/student-template.ts) - Adapt the student template for the data of this group.
                - 🛠️ [`createStudentSheets`](../src/setup/initialize-reports/student-sheets.ts) - Create copies of the student template for each student.
                - 🛠️ [`prepareStatusSheet`](../src/setup/initialize-reports/status.ts) - Prepare the Status sheet to reach to each student sheet.

    - 📄 TODO Report file editing (this is the biggest UI entry point)



📄 (page)
⚡ (zap)
⚒️ (hammer)
🔀 (twisted)
📋 (clipboard)
📊 (bar-chart)



=LET(
  range, E24:G24,

  clean_text, LAMBDA(text,
    LET(
      accent_map, {"Á","A"; "É","E"; "Í","I"; "Ó","O"; "Ú","U"; "Ñ","N"; "Ü","U"},

      no_placeholder, REGEXREPLACE(text, "^\[.*", ""),
      upper_text, UPPER(no_placeholder),

      substitutions, REDUCE(upper_text, SEQUENCE(ROWS(accent_map)), LAMBDA(acc, i, SUBSTITUTE(acc, INDEX(accent_map, i, 1), INDEX(accent_map, i, 2)))),

      ascii_clean, REGEXREPLACE(substitutions, "[^\x00-\x7F]+", ""),
      clean_spaces, TRIM(ascii_clean),
      REGEXREPLACE(clean_spaces, "[,:;\.?!-]+$", "")
    )
  ),

  process_column, LAMBDA(range, prefix,
    LET(
      clean_texts, MAP(range, clean_text), 
      filtered, FILTER(clean_texts, clean_texts <> ""),
      count, IF(ISERROR(filtered), 0, ROWS(filtered)),
      
      IF(count = 0, "", 
        prefix & 
        IF(count = 1, 
          INDEX(filtered, 1), 
          TEXTJOIN("; ", TRUE, CHOOSEROWS(filtered, SEQUENCE(count - 1))) & "; Y " & CHOOSEROWS(filtered, -1)
        ) & "."
      )
    )
  ),

  strenghts, process_column(CHOOSECOLS(range, 1), "DOMINA "),
  weaknesses, process_column(CHOOSECOLS(range, 2), "DEBE FORTALECER "),
  suggestions, process_column(CHOOSECOLS(range, 3), "SE RECOMIENDA "),

  TEXTJOIN(" ", TRUE, strenghts, weaknesses, suggestions)
)