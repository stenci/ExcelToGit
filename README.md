# ExcelToGit
Export Excel workbook macros and data to Git friendly text format.

The macro exports all the VBA code to .bas, .cls and .frm files and all the worksheets to .csv format.

Addins (.xla files) are temporarily set to `.IsAddin = False` in order to select the sheets and export them as .csv.

The macro first exports each sheet as .csv, then saves the file back with its original format.
The side effect is that the next time the file is opened, Excel may not trust it and ask to enable the macros.
(I haven't seen this happening in a long time, I don't know if it's the result of my configuration of some Microsoft 
updates.)

In order for the VBA macro to export the VBA code an Excel setting must be changed:
* In Excel 2003 and earlier: go the Tools - Macros - Security, then click on the Trusted Publishers tab and 
check Trust access to the Visual Basic Project. 
* In Excel 2007 and later: click the Developer item on the main Ribbon, then click the Macro Security item in the 
Code panel, then in the Macro Settings page check the Trust access to the VBA project object model.

### Examples
The first 5 rows are examples use cases, the following row is the real ones that I use for this addin. 
Here is the description of the 5 examples:

#### Addin.xla
The addin is located on the AppData folder and is installed, so it is always loaded. I don't like to keep the git files 
in the AppData folder, and I don't like to uninstall/install the addin every time I need to export to the text files. 
So I configure it with two different folders, and clicking on **Export** will create all the text files 
in **Git folder** and copy the original file from **File to export folder** to **Git folder**. 

#### Macro.xlsm
**Folder** and **Git folder** are the same because this is not an addin, so it is not installed, so it is possible to 
keep it on the same folder with the git files.

#### App
This is not an Excel file. It is just like any other app with its own git rpository which includes two Xl folders, each 
with its own Excel file and its macros. The **Export** button here is missing, because there are no Excel macros or 
sheets to export.

#### App1.xlsm and App2.xlsm
These are two Excel macros, each living in its folder with its exported text files created by clicking on the 
**Export** button. The **Git gui**, **gitk** and **bash** buttons are missing because the git repository that includes 
both the macros is **App**.
