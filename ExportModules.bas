Attribute VB_Name = "ExportModules"
Option Explicit

Global Const COL_EXPORT = 1
Global Const COL_GIT_GUI = 2
Global Const COL_GITK = 3
Global Const COL_GIT_BASH = 4
Global Const COL_NAME = 5
Global Const COL_GIT_FOLDER = 6
Global Const COL_FOLDER = 7

Enum X
  vbext_ct_ActiveXDesigner = 11
  vbext_ct_ClassModule = 2
  vbext_ct_Document = 100
  vbext_ct_MSForm = 3
  vbext_ct_StdModule = 1
End Enum

Sub Export()
  Dim Name As String, GitFolder As String, Folder As String, FullName As String
  Name = ActiveCell.Cells(1, COL_NAME)
  GitFolder = ActiveCell.Cells(1, COL_GIT_FOLDER)
  Folder = ActiveCell.Cells(1, COL_FOLDER)
  FullName = Folder & "\" & Name
  
  If GitFolder = "" Then
    MsgBox "Missing GitFolder", vbCritical
    Exit Sub
  End If
  
  If Dir(GitFolder, vbDirectory) = "" Then
    MsgBox "The GitFolder """ & GitFolder & """ is missing", vbCritical
    Exit Sub
  End If
  
  Dim WB As Workbook
  On Error Resume Next
  Set WB = Workbooks(Name)
  On Error GoTo 0
  
  If WB Is Nothing Then
    MsgBox "Please open the file """ & FullName & """ and try again", vbInformation
    Exit Sub
  End If

  If MsgBox("Export """ & Name & """ to """ & GitFolder & """?", vbYesNo) <> vbYes Then Exit Sub
  
  Dim VBProj
  Set VBProj = WB.VBProject
  
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  
  Dim NewFiles As New Collection
  If Not WB.Saved Then WB.Save
  Shell "cmd /c copy /y """ & FullName & """ """ & GitFolder & """"
  NewFiles.Add Name
  
  Dim OldFiles As New Collection, FName As String
  FName = Dir(GitFolder & "\*")
  Do While FName <> ""
    If LCase(FName) <> ".gitignore" And LCase(FName) <> "readme.md" And LCase(FName) <> "readme.txt" Then OldFiles.Add FName
    FName = Dir()
  Loop
  
  Dim Comp, Components
  Set Components = VBProj.VBComponents
  For Each Comp In Components
    Select Case Comp.Type
      
      Case vbext_ct_ActiveXDesigner
        Stop
      
      Case vbext_ct_ClassModule
        Comp.Export GitFolder & "\" & Comp.Name & ".cls"
        NewFiles.Add Comp.Name & ".cls"
      
      Case vbext_ct_Document
        Comp.Export GitFolder & "\" & Comp.Name & ".cls"
        NewFiles.Add Comp.Name & ".cls"
        
        If Comp.Name <> "ThisWorkbook" Then
          Dim Sh As Worksheet, ShName As String, IsVisible As XlSheetVisibility, ActiveSh As Worksheet
          Set Sh = SheetWithCodeName(WB, Comp.Name)
          IsVisible = Sh.Visible
          ShName = Sh.Name
          If IsAddin(Name) Then WB.IsAddin = False
          If IsVisible <> xlSheetVisible Then Sh.Visible = xlSheetVisible
          Set ActiveSh = WB.ActiveSheet
          WB.Activate
          Sh.Select
          
          ActiveWindow.DisplayFormulas = True
          WB.SaveAs FileName:=GitFolder & "\" & CsvShName(Comp.Name, ShName) & ".csv", FileFormat:=xlCSV, CreateBackup:=False
          ActiveWindow.DisplayFormulas = False
          NewFiles.Add CsvShName(Comp.Name, ShName) & ".csv"
          
          Sh.Name = ShName
          ActiveSh.Activate
          If IsVisible <> xlSheetVisible Then Sh.Visible = IsVisible
          If IsAddin(Name) Then WB.IsAddin = True
          ThisWorkbook.Activate
        End If
        
      Case vbext_ct_MSForm
        Comp.Export GitFolder & "\" & Comp.Name & ".frm"
        NewFiles.Add Comp.Name & ".frm"
        Kill GitFolder & "\" & Comp.Name & ".frx"
      
      Case vbext_ct_StdModule
        Comp.Export GitFolder & "\" & Comp.Name & ".bas"
        NewFiles.Add Comp.Name & ".bas"
      
      Case Else
        Stop
        
    End Select
  Next Comp

  WB.SaveAs FileName:=FullName, FileFormat:=Ext2Format(FullName), CreateBackup:=False
  
  Application.DisplayAlerts = True
  Application.EnableEvents = True
  
  Dim Iold As Integer, Inew As Integer
  For Inew = 1 To NewFiles.Count
    For Iold = 1 To OldFiles.Count
      If LCase(OldFiles(Iold)) = LCase(NewFiles(Inew)) Then
        OldFiles.Remove Iold
        Exit For
      End If
    Next Iold
  Next Inew
  
  Dim Txt As String
  If OldFiles.Count Then
    Txt = "Delete the following files?"
    For Iold = 1 To OldFiles.Count
      Txt = Txt & vbLf & OldFiles(Iold)
    Next Iold
    
    If MsgBox(Txt, vbYesNo) = vbYes Then
      For Iold = 1 To OldFiles.Count
        Kill GitFolder & "\" & OldFiles(Iold)
      Next Iold
    End If
  End If
  
  GoToA2
End Sub

Function SheetWithCodeName(WB As Workbook, CodeName As String) As Worksheet
  For Each SheetWithCodeName In WB.Worksheets
    If SheetWithCodeName.CodeName = CodeName Then Exit Function
  Next SheetWithCodeName
  Set SheetWithCodeName = Nothing
End Function

Function Ext2Format(FileName As String) As XlFileFormat
  If Right(FileName, 4) = ".xla" Then
    Ext2Format = xlAddIn
  ElseIf Right(FileName, 4) = ".xls" Then
    Ext2Format = xlExcel8
  ElseIf Right(FileName, 5) = ".xlsx" Then
    Ext2Format = xlOpenXMLWorkbook
  ElseIf Right(FileName, 5) = ".xlsm" Then
    Ext2Format = xlOpenXMLWorkbookMacroEnabled
  ElseIf Right(FileName, 5) = ".xltm" Then
    Ext2Format = xlOpenXMLTemplateMacroEnabled
  End If
End Function

Function CsvShName(CompName As String, ShName As String) As String
  If CompName = ShName Then
    CsvShName = CompName
  Else
    CsvShName = CompName & " (" & ShName & ")"
  End If
End Function

Sub Refresh()
  Dim WB As Workbook, AI As AddIn
  
  Application.EnableEvents = False
  
  For Each WB In Workbooks
    AddIfMissing WB
  Next WB
  
  For Each AI In AddIns
    If UCase(Right(AI.Name, 4)) <> ".XLL" And UCase(Right(AI.Name, 5)) <> ".XLAM" Then
      AddIfMissing Workbooks(AI.Name)
    End If
  Next AI
  
  Dim C As Integer
  ActiveSheet.UsedRange.EntireColumn.AutoFit
  For C = 1 To ActiveSheet.UsedRange.Columns.Count
    If ActiveSheet.Columns(C).EntireColumn.ColumnWidth > 40 Then ActiveSheet.Columns(C).EntireColumn.ColumnWidth = 40
  Next C
  
  Application.EnableEvents = True
  
  GoToA2
End Sub

Sub GoToA2()
  Application.EnableEvents = False
  Cells(2, 1).Select
  Application.EnableEvents = True
End Sub

Sub AddIfMissing(WB As Workbook)
  Dim R As Integer, DocFolder As String, Name As String
  DocFolder = WB.Path
  Name = WB.Name
  
  For R = 4 To ActiveSheet.UsedRange.Rows.Count
    If Cells(R, COL_NAME) = Name And Cells(R, COL_FOLDER) = DocFolder Then Exit Sub
  Next R
  
  If IsEmpty(Cells(R - 1, 5)) Then R = R - 1
  
  Cells(R, COL_EXPORT) = "Export"
  Cells(R, COL_GIT_GUI) = "Git gui"
  Cells(R, COL_GITK) = "gitk"
  Cells(R, COL_GIT_BASH) = "bash"
  Cells(R, COL_NAME) = Name
  Cells(R, COL_FOLDER) = DocFolder
End Sub

Function IsAddin(Name As String) As Boolean
  IsAddin = UCase(Right(Name, 4)) = ".XLA"
End Function

Sub OpenFolder(FolderName As String)
  If FolderName = "" Then Exit Sub
  If Dir(FolderName, vbDirectory) = "" Then
    MsgBox "Folder """ & FolderName & """ not found", vbCritical
    Exit Sub
  End If
  
  ThisWorkbook.FollowHyperlink FolderName
  
  GoToA2
End Sub

Sub GitGui()
  Dim GitFolder As String
  GitFolder = Cells(ActiveCell.Row, COL_GIT_FOLDER)
  
  If GitFolder = "" Then
    MsgBox "Missing GitFolder", vbCritical
    Exit Sub
  End If
  
  If Dir(GitFolder, vbDirectory) = "" Then
    MsgBox "The GitFolder """ & GitFolder & """ is missing", vbCritical
    Exit Sub
  End If
  
  ChDir2 GitFolder
  Shell """C:\Program Files\Git\cmd\Git-gui.exe"""
  
  GoToA2
End Sub

Sub ChDir2(Path As String)
  If Mid(Path, 2, 1) = ":" Then ChDrive Left(Path, 2)
  ChDir Path
End Sub

Sub Gitk()
  Dim GitFolder As String
  GitFolder = Cells(ActiveCell.Row, COL_GIT_FOLDER)
  
  If GitFolder = "" Then
    MsgBox "Missing GitFolder", vbCritical
    Exit Sub
  End If
  
  If Dir(GitFolder, vbDirectory) = "" Then
    MsgBox "The GitFolder """ & GitFolder & """ is missing", vbCritical
    Exit Sub
  End If
  
  ChDir2 GitFolder
  Shell """C:\Program Files\Git\cmd\Gitk.exe"" --all"
  
  GoToA2
End Sub

Sub GitBash()
  Dim GitFolder As String
  GitFolder = Cells(ActiveCell.Row, COL_GIT_FOLDER)
  
  If GitFolder = "" Then
    MsgBox "Missing GitFolder", vbCritical
    Exit Sub
  End If
  
  If Dir(GitFolder, vbDirectory) = "" Then
    MsgBox "The GitFolder """ & GitFolder & """ is missing", vbCritical
    Exit Sub
  End If
  
  ChDir2 GitFolder
  Shell """C:\Program Files\Git\Git-bash.exe"""
  
  GoToA2
End Sub

Function FolderName(FullPath As String) As String
  FolderName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
End Function
