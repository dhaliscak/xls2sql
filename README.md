# xls2sql
Do you have data in spreadsheet and you need to quickly create INSERT script? Perhaps you will find this tool usefull.

# About app
- C#
- WPF
- .NET 5.0

# Download
Download content of following folder https://github.com/dhaliscak/xls2sql/tree/main/download and run .exe file

# Changelog
## v2.3
- Added Save button to save output query to txt file instead of printing on UI

## v2.2
- fixed: cell value starting with apostrophe
- fixed: empty header value

## v2.1
- added option to drag & file

## v2.0
- added .NET Framework 4.8 project as it by default in Windows (if you don't want to install .NET 5, you can try .NET Framework 4.8 first)

## v1.2
- Removed settings from config file and added them on UI
- Added "First Column" setting
- Improved performance (read of 100k records from 342sec to 5sec)
- Added metric stats to status bar
- Re-sizable window

## v1.1
- new feature: added config file with ability to change batch size (default 1000), prefer nulls (default true), trim whitespaces (default true)
- fixed: replacing "new line" character in name of columns

## v1.0
- initial release
- ability to choose between workbook if spreadsheet contains more of them
- ability to generate script for "CREATE TABLE"
- ability to enter your own database name and table name before generating script

![Thumbnail](https://github.com/dhaliscak/xls2sql/blob/main/download/xls2sql.png)
