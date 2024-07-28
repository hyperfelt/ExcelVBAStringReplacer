# ExcelVBAStringReplacer
A command line tool to replace any string in some VBA code contained in an Excel file.

## Prerequisites

- Microsoft Excel 2007 or higher (the tool uses the `Microsoft.Office.Interop.Excel` library)
- Microsoft .NET Framework v4.7.1 or higher

## Additional notes

If you get an error when you try to process an Excel file containing unprotected VBA code, you need to set up Excel to trust this code. Don't forget to set it back to block unprotected VBA code after you're done!