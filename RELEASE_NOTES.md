# eWorkHelper Release Notes

## Version

v1.2.0

## Release date

2026-08-22

## Overview

This eWorkHelper Excel VSTO Add-in build adds batch unmerge-and-fill while retaining the existing native AutoFilter workflow.

## Features

- Adds an `eWorkHelper` Ribbon tab with a batch-filter command.
- Accepts one text condition per line and removes blank or exact duplicate conditions.
- Supports case-insensitive Equals, Not Equals, Contains, and Not Contains modes.
- Works with Excel tables, existing AutoFilter ranges, and user-selected header rows for ordinary data regions.
- Applies results through the target field's native AutoFilter without changing cell values or creating helper columns or worksheets.
- Clears only the most recently managed target field while leaving other fields' filters in place.
- Adds an About command at the end of the eWorkHelper Ribbon tab, showing a short introduction and the full product version read from assembly metadata.
- Restores eWorkHelper's original batch-filter conditions and match mode when reopening the same active filter; for external Excel filters, loads visible unique target-column values in Equals mode.
- Adds an Unmerge and Fill command that processes every merged area touched by the current selection, including Ctrl-selected areas and partial intersections.
- Uses Excel's native UnMerge operation, deduplicates merge areas, and fills each original area with its value or R1C1 formula without changing worksheet structure.

## Technical information

- C# and .NET Framework 4.8.
- Visual Studio Tools for Office (VSTO) Excel Add-in.
- Excel/Office Interop with embedded interop types.
- Windows Forms dialog and VSTO Ribbon Designer resources.
- Product version (`AssemblyInformationalVersion`): `1.2.0`.
- Assembly and VSTO four-part version: `1.2.0.0`.

## Requirements

- Windows and Microsoft Excel desktop.
- Microsoft Visual Studio Tools for Office Runtime.
- Visual Studio 2022 with VSTO and .NET Framework 4.8 development components to build from source.

## Known limitations

- Excel for the web and Excel for macOS are not supported.
- The current user interface is in Chinese.
- Matching is case-insensitive text matching; regular expressions and formula criteria are not supported.
- Clear-operation state is held in the current Add-in process and is not persisted after the Add-in unloads.
- Excel does not expose this multi-area operation as a transactional VSTO undo unit; if a later area fails, already completed areas cannot be rolled back automatically.
- The compiled component asset is intended for version verification and controlled deployment; it is not a ready-to-install ClickOnce/MSI package.
- Signed deployment and target Office version/bitness combinations require validation in the maintainer's release environment.

## Security and privacy notes

- The public project does not bind to a developer certificate and does not include signing credentials.
- No API keys, tokens, passwords, private keys, certificates, internal endpoints, or local absolute paths are required by the source code.
- Filtering runs locally through the Excel Object Model; the current code does not upload workbook data over a network.
