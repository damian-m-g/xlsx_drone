# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/), featuring Added, Changed, Deprecated,
Removed, Fixed, Security, and others; and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [0.4.0] - 2024-12-09
### Fixed
- #12 (crash with strange format on xlsx file).
- Temporary deploy of any XLSX wasn't being made on the proper folder for Windows (Temp one) but in the cwd.

### Changed
- **zip** library updated from 0.1.32 => 0.3.2.
- **sxmlc** library updated from 4.5.3 => 4.5.4.

### Added
- Dynamically linked lib option for GNU/Linux x64: "libxlsx_drone_x64.so", on the "share" folder.

## [0.3.0] - 2022-11-19
### Changed
- DLLs (on share folder) re-compiled for new version. 

## [0.3.0-pre-2] - 2022-11-19
### Changed
- Updated sxmlc library to version 4.5.3 (2022/09). Contains among other things, a few memory leak fixes.
- Preparing for 0.3.0.

## [0.3.0-pre] - 2022-11-19
### Added
- Support for compatibility with Visual Studio. Code provided by @gvollant on PR (#11).

## [0.2.5] - 2022-11-19
### Fixed
- Issue #5 "Parsing double and locale". Fix provided by @gvollant. On some locales, floating point is different from the
_usual_ dot, like a comma.

## [0.2.4] - 2021-12-05
### Fixed
- Hardcoded relative paths to XLSX internals were working only on Windows. Fixed for other OS.

## [0.2.3] - 2021-11-07
### Fixed
- Added a fallback for when usual environment variable that points to temp directory doesn't exist for the current OS.
According to POSIX/FHS it should be "/tmp" then. Hardcoding that.

## [0.2.2] - 2021-10-21
### Fixed
- Added include for errno.h for non-Windows users.

## [0.2.1] - 2021-10-21
### Fixed
- Issue #2 "tmpnam warning".
- Issue #3 "missing ENVIRONMENT_VARIABLE_TEMP". The line of code that uses this macro was working only on Windows, now
it should work in any other OS too.

### Changed
- Improved README.md.
- zip library updated from 0.1.21 to 0.1.32.

## [0.2.0] - 2020-01-16
### Added
- New function: xlsx_get_last_column().
- New xlsx_sheet_t member: last_column, that won't have value (NULL), unless xlsx_get_last_column() gets called at least 
once.
- New private (static) function, helper of xlsx_get_last_column(): withdraw_alphabetic_chars().
- Several new test assertions for the new functionality.

### Changed
- README.md

## [0.1.9] - 2021-01-10
### Added
- Documentation on the header.
- Several assertions on test_xlsx_load_sheet().
- README.md
- LICENSE

### Changed
- Some variables were renamed to reflect Excel naming.
- Error codes so they are unique.
- xlsx_read_cell() return. Now returns 1 or 0 depending on success or failure.
- Some code style towards simpleness.

## [0.1.8-alpha] - 2021-01-04
### Fixed
- When a cell has no value, but style set, it was triggering problems. Code has been written to take this into account.

## [0.1.7-alpha] - 2021-01-04
### Added
- Unit tests to test empty cells.

## [0.1.6-alpha] - 2021-01-04
### Changed
- Error reporting system. As suggested by the web, I'm not using anymore errno error system to give info about an error
since that system is supposed to be used only by the OS. Now I'm using a custom one setting xlsx_errno static global
variable, accessible through xlsx_get_xlsx_errno().
  
### Added
- Unit test regarding error reporting system change.

### Fixed
- Manually fixed problems with the zip library (since were acknowledged by the maker but didn't release a new version).
The function that opens and deploy the zip weren't returning a negative number if fails under certain circumstances.

## [0.1.5-alpha] - 2020-12-31
### Added
- Added a few tests for Unicode (UTF-8) support.
  
### Updated
- Libraries used: zip 0.1.20 -> 0.1.21; sxmlc 4.3.0 -> 4.5.1.

## [0.1.4-alpha] - 2020-12-29
### Fixed
- Explicit typecast made on several long values that have been assigned to int holders.
- A string cell value could have a style associated, this wasn't taken into account before. Code was written to embrace 
this possibility.
- set_cell_data_values_for_number() wasn't working perfect. If an "E" is present in a cell value that isn't a string,
that values are always returned as a double, even if it isn't a floating point value.
  
### Added
- Finished test_xlsx_read_cell(). A total of 662 assertions were written for unit testing. All passing at this version.

## [0.1.3-alpha] - 2020-12-27
### Added
- Finished test_xlsx_load_sheet() & test_xlsx_unload_sheet() & test_xlsx_close().
- Although the user shouldn't want to access freed data after it was freed, he might want to give it a try. I wasn't
setting pointers to freed data to NULL, and now I decided to do it. These changes were implemented in
xlsx_unload_sheet() & xlsx_close().

## [0.1.2-alpha] - 2020-12-26
### Added
- Finished test_xlsx_open() assertions battery of tests.

## [0.1.1-alpha] - 2020-12-24
### Fixed
- Problem in the creation of the temporary directory where the XLSX is deployed.

### Changed
- The way data related to default styles is loaded. Previously, data were loaded on runtime and also had to be freed 
  on exit. Now this data is compiled, and doesn't need to be freed.

## [0.1.0-alpha] - 2020-12-22
### Fixed
- Shared strings could not exist when the XLSX is empty or has no strings. The code was modified to contemplate this 
  possibility.