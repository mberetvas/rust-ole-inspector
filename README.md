# rust-ole-inspector

A **Rust CLI** for **Windows** that discovers **COM objects** and checks their programmatic usability.

## Overview

This tool iterates through the **Windows Registry** at `HKEY_CLASSES_ROOT\CLSID` to map **Class IDs (CLSIDs)** to their human-readable **Program IDs (ProgIDs)**, using the official `windows` crate. It provides insights into which COM objects are available on your system and how easily they can be used programmatically.

## Features

- 🔍 **Registry Scanning**: Enumerates all COM objects registered in the Windows Registry
- 🔗 **CLSID to ProgID Mapping**: Automatically maps Class IDs to their Program IDs
- 📊 **Usability Assessment**: Evaluates how programmatically usable each COM object is
- 🏗️ **Multi-Architecture Support**: Handles both 32-bit and 64-bit registry views
- 🔐 **Privilege Detection**: Warns when not running with elevated privileges
- 🎯 **Filtering**: Search for specific COM objects by name or description
- 📝 **Detailed Output**: Optional verbose mode for complete information

## Requirements

- **Windows Operating System**: This tool is Windows-only as it accesses the Windows Registry
- **Rust**: Latest stable version
- **Elevated Privileges**: Administrator rights recommended for a complete scan

## Installation

```bash
# Clone the repository
git clone https://github.com/mberetvas/rust-ole-inspector.git
cd rust-ole-inspector

# Build the project
cargo build --release

# The executable will be in target/release/rust-ole-inspector.exe
```

## Usage

### Basic Scan

```bash
# Scan both 32-bit and 64-bit registry views (default)
rust-ole-inspector.exe
```

### Run with Administrator Privileges

For a complete scan, run from an elevated command prompt:

```bash
# Right-click Command Prompt/PowerShell and select "Run as Administrator"
rust-ole-inspector.exe
```

### Advanced Options

```bash
# Show detailed information for each COM object
rust-ole-inspector.exe --verbose

# Scan only 32-bit registry view
rust-ole-inspector.exe --scan-32bit

# Scan only 64-bit registry view
rust-ole-inspector.exe --scan-64bit

# Limit results to first 100 objects
rust-ole-inspector.exe --limit 100

# Filter by ProgID or description
rust-ole-inspector.exe --filter "Excel"

# Combine options
rust-ole-inspector.exe --verbose --filter "Word" --limit 10
```

### Command-Line Options

- `-v, --verbose`: Show detailed information for each COM object
- `--scan-32bit`: Scan 32-bit registry view
- `--scan-64bit`: Scan 64-bit registry view (default on 64-bit systems)
- `-l, --limit <NUMBER>`: Limit the number of results (0 = no limit)
- `-f, --filter <TEXT>`: Filter by ProgID substring (case-insensitive)
- `-h, --help`: Display help information
- `-V, --version`: Display version information

## Output Explanation

The tool provides:

1. **Privilege Status**: Whether running with elevated privileges
2. **Scan Results**: Number of COM objects found in each registry view
3. **Statistics**: Total objects found and percentage with ProgIDs
4. **Usability Rating**: Assessment of how programmatically usable each COM object is:
   - ✓ **High**: Has both ProgID and description (easily usable)
   - ~ **Medium**: Has ProgID but no description (usable by name)
   - ~ **Low**: Has description but no ProgID (requires CLSID)
   - ✗ **Very Low**: No ProgID or description (requires CLSID, poorly documented)

## Example Output

```
=== Windows COM Object Inspector ===

✓ Running with elevated privileges

Scanning 32-bit registry view...
Found 1234 COM objects in 32-bit view

Scanning 64-bit registry view...
Found 1456 COM objects in 64-bit view

=== Results ===
Total unique COM objects found: 2345

COM objects with ProgID: 987 (42.1%)
COM objects without ProgID: 1358

--- COM Objects with ProgID ---
  Excel.Application ({00024500-0000-0000-C000-000000000046})
  Word.Application ({000209FF-0000-0000-C000-000000000046})
  Shell.Application ({13709620-C279-11CE-A49E-444553540000})
  ...
```

## Technical Details

### Registry Views

Windows maintains separate registry views for 32-bit and 64-bit applications:
- **32-bit view** (`KEY_WOW64_32KEY`): Contains COM objects for 32-bit applications
- **64-bit view** (`KEY_WOW64_64KEY`): Contains COM objects for 64-bit applications

By default, this tool scans both views to provide a complete picture.

### Programmatic Usability

COM objects can be instantiated in several ways:
1. **By ProgID**: `CreateObject("Excel.Application")` - Most user-friendly
2. **By CLSID**: `CoCreateInstance({...})` - More verbose, requires GUID

Objects with ProgIDs are considered more programmatically usable as they can be referenced by a human-readable name.

## Development

```bash
# Run in development mode
cargo run

# Run with arguments
cargo run -- --verbose --limit 10

# Run tests (if available)
cargo test

# Format code
cargo fmt

# Run linter
cargo clippy
```

## License

See [LICENSE](LICENSE) file for details.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

