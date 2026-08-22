# eWorkHelper Release Notes

## Version

v1.0.0

## Release date

2026-08-22

## Overview

This is the first public source release of eWorkHelper, an Excel VSTO Add-in that provides a batch text-filtering workflow backed by Excel's native AutoFilter.

## Features

- Adds an `eWorkHelper` Ribbon tab with a batch-filter command.
- Accepts one text condition per line and removes blank or exact duplicate conditions.
- Supports case-insensitive Equals, Not Equals, Contains, and Not Contains modes.
- Works with Excel tables, existing AutoFilter ranges, and user-selected header rows for ordinary data regions.
- Applies results through the target field's native AutoFilter without changing cell values or creating helper columns or worksheets.
- Clears only the most recently managed target field while leaving other fields' filters in place.

## Technical information

- C# and .NET Framework 4.8.
- Visual Studio Tools for Office (VSTO) Excel Add-in.
- Excel/Office Interop with embedded interop types.
- Windows Forms dialog and VSTO Ribbon Designer resources.
- Assembly version: `1.0.0.0`.

## Requirements

- Windows and Microsoft Excel desktop.
- Microsoft Visual Studio Tools for Office Runtime.
- Visual Studio 2022 with VSTO and .NET Framework 4.8 development components to build from source.

## Known limitations

- Excel for the web and Excel for macOS are not supported.
- The current user interface is in Chinese.
- Matching is case-insensitive text matching; regular expressions and formula criteria are not supported.
- Clear-operation state is held in the current Add-in process and is not persisted after the Add-in unloads.
- This source release does not include a signed installer or prebuilt deployment package.
- Signed deployment and target Office version/bitness combinations require validation in the maintainer's release environment.

## Security and privacy notes

- The public project does not bind to a developer certificate and does not include signing credentials.
- No API keys, tokens, passwords, private keys, certificates, internal endpoints, or local absolute paths are required by the source code.
- Filtering runs locally through the Excel Object Model; the current code does not upload workbook data over a network.
