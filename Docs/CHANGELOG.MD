# Changelog

All notable changes to this project will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.0.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## [Unreleased]

### Changed
- **Refactored monolithic main.rs into modular architecture**: Split 609-line single file into 7 focused modules:
  - `types.rs`: CLI arguments and ComObject data structures
  - `registry.rs`: Windows Registry operations (CLSID enumeration, ProgID/description lookup)
  - `filter.rs`: Filtering logic supporting multiple filter types
  - `display.rs`: Result presentation and export functionality (console, CSV, TXT)
  - `console.rs`: Console utilities (UTF-8 setup, header art)
  - `security.rs`: Privilege detection and warnings
  - `main.rs`: Clean orchestration layer
  - **TLDR**: Improved maintainability, testability, and code clarity while maintaining all existing functionality.

## [0.1.0] - YYYY-MM-DD

### Added
- Initial release of the Rust OLE Inspector tool.
- Basic functionality for inspecting OLE files.