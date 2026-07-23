# Sarwar Smart PCMS

Original 32-bit Microsoft Excel project-controls add-in built with Excel-DNA and C#.

## Current implemented functions
- Automatic Excel Ribbon
- WBS coloring, grouping and summary
- Monthly Gantt drawing
- Unit distribution and S-curve
- Primavera XER table import and explorer
- Calendar extraction
- Basic XER export
- Schedule health check
- XER revision comparison
- Cell tools
- Dashboard worksheet
- Automatic installation through Excel XLSTART

## Development status
This package is a working source-code foundation, not a claim of feature-for-feature parity with any commercial product. Advanced buttons such as flow path, out-of-sequence repair, half-step XER and cleanup are visible but marked as staged foundations until they are validated against real schedules.

## Build without Visual Studio on your computer
Use GitHub Actions. See `docs/A-Z-GITHUB-GUIDE.md`.

## Excel compatibility
The project is configured for **32-bit Excel (x86)**. A separate x64 build should be created for 64-bit Excel.
