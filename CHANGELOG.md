# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

## [0.1.5-alpha] - 2020-12-31
### Added
- Added a few tests for unicode (UTF-8) support.
  
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
setting pointers to freed data to NULL and now I decided to do it. This changes were implemented in xlsx_unload_sheet()
& xlsx_close().

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