# eWorkHelper Release Notes

> **权威发布历史已统一到 [`docs/CHANGELOG.md`](docs/CHANGELOG.md)。**
> 本文件只保留当前版本的对外摘要；历史版本条目、日期与修订详情一律以 `docs/CHANGELOG.md` 为准，
> 请勿在本文件中单独新增版本条目，以免再次产生两条并行发布历史。

## Version

v1.2.0

## Release date

2026-08-22

## Overview

This eWorkHelper Excel VSTO Add-in build adds batch unmerge-and-fill while retaining the existing native AutoFilter workflow.

## Features

- Adds a custom `工作助手` (Work Assistant) Ribbon tab with a batch-filter command.
- Accepts one text condition per line and removes blank or exact duplicate conditions, comparing duplicates case-insensitively.
- Supports case-insensitive Equals, Not Equals, Contains, and Not Contains modes.
- Works with Excel tables, existing AutoFilter ranges, and user-selected header rows for ordinary data regions.
- Applies results through the target field's native AutoFilter without changing cell values or creating helper columns or worksheets.
- Clears only the most recently managed target field while leaving other fields' filters in place.
- Matches on the cell's displayed text, so date-, percentage-, and currency-formatted columns filter the same way they look on screen.
- Adds an About command at the end of the Ribbon tab, showing a short introduction and the full product version read from assembly metadata.
- Restores eWorkHelper's original batch-filter conditions and match mode when reopening the same active filter; for external Excel filters, loads visible unique target-column values in Equals mode.
- Reports long-running progress on the Excel status bar and lets the user cancel with Esc; cursor and Excel application state are restored on every exit path.
- Adds an Unmerge and Fill command that processes every merged area touched by the current selection, including Ctrl-selected areas and partial intersections.
- Uses Excel's native UnMerge operation, deduplicates merge areas, and fills each original area with its value or R1C1 formula without changing worksheet structure.
- Skips error-valued cells, runs with manual calculation, and reports clear messages for array formulas (CSE) and protected sheets.
- Releases the Excel COM objects it creates deterministically, so workbooks are no longer kept alive after the user closes them.

## Technical information

- C# and .NET Framework 4.8.
- Visual Studio Tools for Office (VSTO) Excel Add-in.
- Excel/Office Interop with embedded interop types.
- Windows Forms dialog and VSTO Ribbon Designer resources.
- Product version (`AssemblyInformationalVersion`): `1.2.0`.
- Assembly and VSTO four-part version: `1.2.0.0`.
- Manifest signing is mandatory; supply the certificate thumbprint with `/p:ManifestCertificateThumbprint=<thumbprint>`
  or the `IWORKHELPER_MANIFEST_CERT_THUMBPRINT` environment variable. See [`README.md`](README.md#编译).

## Requirements

- Windows and Microsoft Excel desktop.
- Microsoft Visual Studio Tools for Office Runtime.
- Visual Studio 2022 with VSTO and .NET Framework 4.8 development components to build from source.
- A code-signing certificate in the build machine's certificate store to sign the VSTO manifest.

## Known limitations

- Excel for the web and Excel for macOS are not supported.
- The current user interface is in Chinese.
- Matching is case-insensitive text matching; regular expressions and formula criteria are not supported.
- Case-insensitive matching uses .NET `OrdinalIgnoreCase`, which is not a full Unicode case fold: locale-specific pairs such as Turkish/Azerbaijani `i`/`İ` are not treated as the same character.
- A condition may be at most 8192 characters, and the match set must stay within Excel's filter value-list capacity (10000 items); the tool reports a clear message rather than hiding every row or surfacing a raw COM error.
- Clear-operation state is held in add-in process memory and is scoped to the current Excel window (one Ribbon instance per window); other windows do not see it and it is not persisted after the add-in unloads.
- Excel does not expose this multi-area operation as a transactional VSTO undo unit; if a later area fails, already completed areas cannot be rolled back automatically. The error message reports how many areas were completed.
- The compiled component asset is intended for version verification and controlled deployment; it is not a ready-to-install ClickOnce/MSI package.
- Signed deployment and target Office version/bitness combinations require validation in the maintainer's release environment.

## Security and privacy notes

- The public project does not bind to a developer certificate and does not include signing credentials.
- No API keys, tokens, passwords, private keys, certificates, internal endpoints, or local absolute paths are required by the source code.
- Filtering runs locally through the Excel Object Model; the current code does not upload workbook data over a network.
