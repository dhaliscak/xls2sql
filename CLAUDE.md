# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

WPF desktop tool that converts Excel/CSV files into SQL INSERT (and optionally CREATE TABLE) scripts. Two parallel projects targeting different runtimes — both implement the same feature set.

## Projects

| Project | Target | Notes |
|---------|--------|-------|
| `xls2sql/` | .NET 5.0 (net5.0-windows) | Primary project |
| `xls2sql48/` | .NET Framework 4.8 | Alternate for users without .NET 5 installed |

Both use `ExcelDataReader` + `ExcelDataReader.DataSet` NuGet packages. The .NET 4.8 project uses `packages/` folder (packages.config); the .NET 5 project uses SDK-style PackageReference.

## Build

Open `xls2sql.sln` in Visual Studio. No CLI build tooling is configured — build via VS or `msbuild`.

Build from CLI (if msbuild available):
```
msbuild xls2sql.sln /p:Configuration=Release
```

Publish self-contained exe (net5.0 project):
```
dotnet publish xls2sql/xls2sql.csproj -c Release -r win-x64 --self-contained
```

## Architecture

Both projects are single-window WPF apps with no MVVM — all logic lives in `MainWindow.xaml.cs`.

**Flow:**
1. User opens/drops a file → `LoadFile()` reads via `ExcelReaderFactory`, caches result in `_cachedDataSet` field
2. If multiple sheets, workbook `ComboBox` appears for selection
3. "Generate"/"Save" → `ReadSettings()` snapshots UI state into `SqlSettings`, then `GenerateSQLQuery(SqlSettings)` runs on a background thread via `Task.Run`
4. "Save" writes generated SQL to `{tableName}.txt` next to the `.exe`

**Key methods in `MainWindow.xaml.cs`:**
- `LoadFile(string)` — reads file, populates `_cachedDataSet`, updates workbook ComboBox
- `ReadSettings()` — captures all UI control values into a `SqlSettings` object (allows off-thread execution)
- `GenerateSQLQuery(SqlSettings)` — static; core SQL generation; streams rows directly into `StringBuilder` without intermediate buffering; batch-splits INSERTs at `Separator` row count
- `GetColumnNames()` — first row treated as headers; empty headers excluded
- `GenerateColumnNamesForCreateTable()` — optionally prepends `Id` column (INT IDENTITY or UNIQUEIDENTIFIER)

**`SqlSettings`** (bottom of `MainWindow.xaml.cs`) — plain data class holding all generation parameters including a `DataTable` reference captured before `Task.Run`.

**SQL output dialect:** T-SQL (SQL Server). Uses `USE {db}`, `SET ANSI_NULLS ON`, bracketed column names `[col]`, `varchar(max)` for all columns.

**Cell value escaping:** single quotes escaped as `''`. NULL preference: empty string → `NULL` when "Prefer Nulls" checked. Note: the `N'` Unicode prefix is only applied from the second column onward (artifact of `string.Join("', N'", ...)` — intentional behavior to preserve).

**`xls2sql48/ExcelHelper.cs`** — dead code; not used by `MainWindow`. Was an OleDb-based reader from an earlier version.

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
| `txtEditor` | Output SQL text |
| `txtStatus` | Timing/status bar |
