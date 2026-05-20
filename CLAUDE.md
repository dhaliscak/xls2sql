# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Overview

WPF desktop tool that converts Excel/CSV files into SQL INSERT (and optionally CREATE TABLE) scripts. Single project (`xls2sql/xls2sql.csproj`) with multi-targeting for both .NET 5 and .NET Framework 4.8.

## Build

Open `xls2sql.sln` in Visual Studio, or from CLI:

```
msbuild xls2sql.sln /p:Configuration=Release
```

Publish self-contained exe (net5.0, win-x64):
```
dotnet publish xls2sql/xls2sql.csproj -c Release -r win-x64 --self-contained
```

Both targets use SDK-style `PackageReference` with `ExcelDataReader` + `ExcelDataReader.DataSet` v3.6.0.

## Architecture

Single-window WPF app with no MVVM — all logic lives in `xls2sql/MainWindow.xaml.cs`.

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
