# General code
* [ ] Update `transferData` function to grow only on rows and spill over content.

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

# Menu
## Assistance
* [ ] Recalculate periods

## Persistant Data
* [ ] Normalize weights (for grades and subjects) 
* [ ] Change weights