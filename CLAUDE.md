# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

WPF desktop tool that converts Excel/CSV files into SQL INSERT (and optionally CREATE TABLE) scripts. Single project (`xls2sql/xls2sql.csproj`) multi-targeting .NET 8 (`net8.0-windows`) and .NET Framework 4.8.

## Build

Open `xls2sql.sln` in Visual Studio, or from CLI:

```
dotnet build xls2sql/xls2sql.csproj -c Release
```

Publish self-contained exe (net8.0, win-x64):
```
dotnet publish xls2sql/xls2sql.csproj -c Release -f net8.0-windows -r win-x64 --self-contained
```

Packages: `Sylvan.Data.Excel` (.xls/.xlsx/.xlsb), `Sylvan.Data.Csv` (.csv). No editor package: output is shown by the custom `OutputViewer`. `LangVersion` is `latest` so net48 can use C# 8+ syntax. Code must compile for net48 too (no span-only APIs without `#if`).

## Architecture

Single-window WPF app, no MVVM. UI glue in `xls2sql/MainWindow.xaml.cs`; WPF-free reading/generation in `xls2sql/SqlGenerator.cs`.

**Flow:**
1. Open/drop file → `LoadFile()` reads sheet names only (`SheetReader.ReadSheetNames`), clears `_sheetCache`
2. Sheet selected (auto when single sheet) → `CmbWorkbook_SelectionChanged` starts background `SheetReader.Load` into `_sheetCache` (`Dictionary<sheetName, Task<SheetData>>`, UI-thread only). File is parsed once per sheet; later edits to the file on disk are not seen until reload (same as the original DataSet cache)
3. Generate/Save → `ReadSettings()` snapshots UI into `SqlSettings`; `GenerateSQLQueryAsync` awaits the cached sheet, then `SqlGenerator.Generate` on the thread pool
4. Generate first drops the previous `_output` (old and new are never alive together), builds `OutputIndex` off the UI thread, calls `txtEditor.SetOutput`, then `ReleaseFreedMemory()` (LOH compaction; `GCCollectionMode.Aggressive` on .NET 7+ to decommit). Save streams `SqlOutput` chunks to `{tableName}.txt` next to the `.exe` (same bytes as `File.WriteAllText`). Copy All decodes `_output` to one string only on click

**Theming (`xls2sql/ThemeManager.cs`, `xls2sql/Themes/`):**
- `Themes/Light.xaml` / `Themes/Dark.xaml` — palettes with identical `Theme.*` brush keys (add a key to both). Merged into `App.xaml` at index 0; `ThemeManager` swaps that entry at runtime
- `Themes/Controls.xaml` — implicit styles for Label, Button, TextBox, CheckBox, ComboBox/ComboBoxItem, ScrollBar, ContextMenu/MenuItem, ToolTip, ResizeGrip, plus keyed `IconButton`. Full replacement templates because Aero2 hardcodes light colors. All colors via `DynamicResource Theme.*`; never hardcode colors or use `SystemColors` in XAML/code
- `ThemeManager` — `ThemeMode` System/Light/Dark, `Cycle()` System→Light→Dark→System. System reads `HKCU\...\Themes\Personalize\AppsUseLightTheme` and follows live changes via `SystemEvents.UserPreferenceChanged`. Mode is **not persisted** (each start = System; do not add settings storage). `ApplyTitleBar` sets `DWMWA_USE_IMMERSIVE_DARK_MODE` (20, fallback 19). `Initialize`/`Shutdown` from `App.OnStartup`/`OnExit`
- `OutputViewer` draws itself: colors come from its `Background`/`Foreground`/`SelectionBrush` dependency properties, bound to `Theme.Output*` in `MainWindow.xaml`

**`SqlGenerator.cs`:**
- `SheetData` — all rows as `string[]` (row 0 = header); each row has the reader's `FieldCount` cells, padded with `""`
- `SheetReader` — Sylvan Excel with `ExcelSchema.NoHeaders`; Sylvan CSV with `HasHeaders = false`, Windows-1250 fallback encoding, auto-detected delimiter
- `SqlGenerator.Generate` → `SqlOutput`: data rows formatted in parallel chunks (`ChunkRows`) by per-thread `ChunkWorker`s (reused `StringBuilder`; on .NET Core encoded straight to UTF-8 via `Encoder` without a temp string). Result is never concatenated. Batch split (`;` + new INSERT head) depends only on the global data-row index, so chunking does not change output
- `SqlOutput` — ordered UTF-8 `byte[]` chunks; half the memory of a UTF-16 string, no 2 GB object limit. Chunk boundaries never split a character
- `RowFormatter` — allocation-free cell escaping/trim/NULL handling. Its doc comment states the original string-based logic it must match exactly
- `SqlSettings` — plain settings DTO

**Output must stay byte-identical** to the original algorithm: `value.Replace("'", "''")` → optional `Trim()` → empty→`"NULL"` if Prefer Nulls → `('v0', N'v1', ...)` → `Replace("N'NULL'", "NULL")` when Prefer Nulls. Consequences to preserve: no `N'` on column 0; column 0 empty→`'NULL'` (quoted); column 1+ literal `NULL` text also becomes bare `NULL`; value starting `NULL'` loses its `N'` prefix. Row width = `min(headerCount, row cells)` positional, even when empty headers were skipped.

**Cell text:** via `GetString`. Excel dates come out ISO 8601 (`2024-03-07T13:05:00`, or `2024-12-31` for date-only).

**SQL output dialect:** T-SQL (SQL Server). `USE {db}`, `SET ANSI_NULLS ON`, `[col]` names, `varchar(max)` columns, optional `Id` first column (INT IDENTITY or UNIQUEIDENTIFIER NEWSEQUENTIALID). Line breaks are `Environment.NewLine`; each INSERT batch is a single line.

## UI Controls (MainWindow.xaml)

| Name | Purpose |
|------|---------|
| `txtDatabaseName` | Database name in generated `USE` statement |
| `txtTableName` | Table name in INSERT/CREATE TABLE |
| `ckbCreateTable` | Toggle CREATE TABLE generation |
| `cmbFirstColumn` | Add auto-id first column (None / INT IDENTITY / NEWSEQUENTIALID) |
| `txtSeparator` | Rows per INSERT batch (default 1000) |
| `ckbTrimWhiteSpaces` | Trim cell values |
| `ckbPrefferNulls` | Treat empty strings as NULL |
| `txtFilepath` | Read-only; set via Open File or drag-and-drop |
| `cmbWorkbook` | Sheet selector; hidden when file has only one sheet |
| `txtEditor` | `OutputViewer` (in `OutputViewer.cs`) inside a `ScrollViewer` with `CanContentScroll=True`. Read-only, virtualized: hard-wraps lines into fixed Consolas rows (columns counted in UTF-8 bytes, like xls2sql-rs `viewer.rs`) and decodes/draws only visible rows. Mouse selection, Ctrl+C, Ctrl+A, context menu. Replaced TextBox (~25s layout for 30M chars) and AvalonEdit (lays out each whole ~600K-char INSERT line; 2.8 GB + 64s to scroll a 239 MB script) |
| `btnCopyAll` | Next to "Output" label; copies full output to clipboard. Hidden until Generate produces output (mirrors xls2sql-rs) |
| `txtStatus` | Timing/status bar |
| `btnTheme` | Top-right icon button (`IconButton` style, overlays the window; first row has right margin for it). Cycles theme; glyph E7F4 monitor / E706 sun / E708 moon (Segoe Fluent Icons / MDL2 Assets), tooltip shows current mode |

## Release

Version lives in `xls2sql.csproj` `<Version>` and the `MainWindow.xaml` window title — bump both. Publish each target (framework-dependent, not self-contained) into `download/.net 8.0/vX.Y/` and `download/.net 4.8/vX.Y/`; keep older version folders. Add a `## vX.Y` changelog entry at the top of README.md and point its Download section at the new folders.
