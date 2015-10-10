# ExcelToGit
Export Excel workbook macros and data to Git friendly text format.

The macro exports all the VBA code to .bas, .cls and .frm files and all the worksheets to .csv format.

Add-ins (.xla files) are temporarily set to `.IsAddin = False` in order to select the sheets and export them as .csv.

The macro first exports each sheet as .csv, then saves the file back with its original format. 
The side effect is that the next time the file is opened, Excel will not trust it and ask to enable the macros.

In order for the VBA macro to export the VBA code a setting must be changed:
* In Excel 2003 and earlier: go the Tools - Macros - Security, then click on the Trusted Publishers tab and 
check Trust access to the Visual Basic Project. 
* In Excel 2007 and later: click the Developer item on the main Ribbon, then click the Macro Security item in the 
Code panel, then in the Macro Settings page check the Trust access to the VBA project object model.
