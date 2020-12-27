# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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