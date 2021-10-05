# xls2sql
Do you have data in spreadsheet and you need to quickly create INSERT script? Perhaps you will find this tool usefull.

# About app
- C#
- WPF
- .NET 5.0

# Download
Download content of following folder https://github.com/dhaliscak/xls2sql/tree/main/download and run .exe file

# Changelog

## v1.1
- new feature: added configfile with ability to change batch size (default 1000), prefer nulls (default true), trim whitespaces (default true)
- fixed: replacing "new line" character in name of columns

## v1.0
- initial release
- ability to choose between workbook if spreadsheet contains more of them
- ability to generate script for "CREATE TABLE"
- ability to enter your own database name and table name before generating script

![Thumbnail](https://github.com/dhaliscak/xls2sql/blob/main/download/xls2sql.png)