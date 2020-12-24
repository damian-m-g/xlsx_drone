# Changelog
All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

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