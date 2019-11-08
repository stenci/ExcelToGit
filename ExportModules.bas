Attribute VB_Name = "ExportModules"
Option Explicit

Public Declare PtrSafe Function SetCurrentDirectoryA Lib "kernel32" (ByVal lpPathName As String) As Long

Global Const COL_EXPORT = 1
Global Const COL_GIT_GUI = 2
Global Const COL_GITK = 3
Global Const COL_GIT_BASH = 4
Global Const COL_NAME = 5
Global Const COL_FILE_FOLDER = 6
Global Const COL_GIT_FOLDER = 7

Enum X
  vbext_ct_ActiveXDesigner = 11
  vbext_ct_ClassModule = 2
  vbext_ct_Document = 100
  vbext_ct_MSForm = 3
  vbext_ct_StdModule = 1
End Enum

Sub Export()
  Dim Name As String, GitFolder As String, FileFolder As String, FullName As String
  Name = ActiveCell.Cells(1, COL_NAME)
  GitFolder = ActiveCell.Cells(1, COL_GIT_FOLDER)
  FileFolder = ActiveCell.Cells(1, COL_FILE_FOLDER)
  FullName = FileFolder & "\" & Name
  
  GoToNameColumn
  
  If GitFolder = "" Then
    MsgBox "Missing GitFolder", vbCritical
    Exit Sub
  End If
  
  If Dir(GitFolder, vbDirectory) = "" Then
    MsgBox "The GitFolder """ & GitFolder & """ is missing", vbCritical
    Exit Sub
  End If
  
  If MsgBox("Export """ & Name & """ to """ & GitFolder & """?", vbYesNo) <> vbYes Then Exit Sub
  
  Dim WB As Workbook
  On Error Resume Next
  Set WB = Workbooks(FullName)
  On Error GoTo 0
  If WB Is Nothing Then
    Application.EnableEvents = False
    Set WB = Workbooks.Open(FullName, UpdateLinks:=False, Editable:=True)
    Application.EnableEvents = True
  End If
  
  Dim VBProj
  Set VBProj = WB.VBProject 'see https://github.com/stenci/ExcelToGit
  
  Application.EnableEvents = False
  Application.DisplayAlerts = False
  
  Dim NewFiles As New Collection
  If WB.Path <> GitFolder Then ExecuteCommand "copy /y """ & FullName & """ """ & GitFolder & """"
  NewFiles.Add WB.Name
  
  Dim OldFiles As New Collection, FName As String
  FName = Dir(GitFolder & "\*")
  Do While FName <> ""
    If LCase(FName) <> ".gitignore" And _
       LCase(FName) <> ".gitattributes" And _
       LCase(FName) <> "readme.md" And _
       LCase(FName) <> "readme.txt" _
       Then OldFiles.Add FName
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
          If IsAddin(WB.Name) Then WB.IsAddin = False
          Set ActiveSh = WB.ActiveSheet
          If IsVisible <> xlSheetVisible Then Sh.Visible = xlSheetVisible
          WB.Activate
          Sh.Select
          
          ActiveWindow.DisplayFormulas = True
          WB.SaveAs FileName:=GitFolder & "\" & CsvShName(Comp.Name, ShName) & ".csv", FileFormat:=xlCSV, CreateBackup:=False
          ActiveWindow.DisplayFormulas = False
          NewFiles.Add CsvShName(Comp.Name, ShName) & ".csv"
          
          Sh.Name = ShName
          ActiveSh.Activate
          If IsVisible <> xlSheetVisible Then Sh.Visible = IsVisible
          If IsAddin(WB.Name) Then WB.IsAddin = True
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

  If WB Is ThisWorkbook Then
    WB.SaveAs FileName:=FullName, FileFormat:=Ext2Format(FullName), CreateBackup:=False
  Else
    WB.Close
  End If
  
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
End Sub

Function SheetWithCodeName(WB As Workbook, CodeName As String) As Worksheet
  For Each SheetWithCodeName In WB.Worksheets
    If UCase(SheetWithCodeName.CodeName) = UCase(CodeName) Then Exit Function
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
  
  GoToNameColumn
End Sub

Sub GoToNameColumn()
  Application.EnableEvents = False
  Cells(ActiveCell.Row, 5).Select
  Application.EnableEvents = True
End Sub

Sub AddIfMissing(WB As Workbook)
  Dim R As Integer, DocFolder As String, Name As String
  DocFolder = WB.Path
  Name = WB.Name
  
  For R = 4 To ActiveSheet.UsedRange.Rows.Count
    If Cells(R, COL_NAME) = Name And Cells(R, COL_FILE_FOLDER) = DocFolder Then Exit Sub
  Next R
  
  If IsEmpty(Cells(R - 1, 5)) Then R = R - 1
  
  Cells(R, COL_EXPORT) = "Export"
  Cells(R, COL_GIT_GUI) = "Git gui"
  Cells(R, COL_GITK) = "gitk"
  Cells(R, COL_GIT_BASH) = "bash"
  Cells(R, COL_NAME) = Name
  Cells(R, COL_FILE_FOLDER) = DocFolder
End Sub

Function IsAddin(Name As String) As Boolean
  IsAddin = UCase(Right(Name, 4)) = ".XLA"
End Function

Sub OpenFolder(FolderName As Range)
  If FolderName = "" Then Exit Sub
  If Dir(FolderName, vbDirectory) = "" Then
    MsgBox "Folder """ & FolderName & """ not found", vbCritical
    Exit Sub
  End If
  
  Dim FileName As String
  FileName = FolderName.Value & "\" & FolderName.Worksheet.Cells(FolderName.Row, COL_NAME)
  
  Dim SystemRoot As String, Shell As New Shell
  SystemRoot = Environ("SystemRoot")
  If Dir(FileName) <> "" Then
    Shell.ShellExecute SystemRoot & "\Explorer.exe", "/select, """ & FileName & """", "", "open", 1
  Else
    Shell.ShellExecute SystemRoot & "\Explorer.exe", """" & FolderName.Value & """", "", "open", 1
  End If
  
  GoToNameColumn
End Sub

Sub ExecuteCommand(Command As String)
  Static LastId As Integer
  Dim CmdFile As String
  LastId = LastId + 1
  CmdFile = Environ("TEMP") & "\ExcelToGit" & LastId & ".cmd"
  Open CmdFile For Output As #1
  Print #1, Replace(Command, "%", "%%")
  Close #1
  Shell CmdFile
End Sub

Function GitExeFolder() As String
  If Dir("C:\Program Files (x86)\Git", vbDirectory) <> "" Then
    GitExeFolder = "C:\Program Files (x86)\Git"
  ElseIf Dir("C:\Program Files\Git", vbDirectory) <> "" Then
    GitExeFolder = "C:\Program Files\Git"
  Else
    Stop
  End If
End Function

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
  Shell """" & GitExeFolder & "\cmd\Git-gui.exe"""
  
  GoToNameColumn
End Sub

Sub ChDir2(Path As String)
  If Left(Path, 2) = "\\" Then
    SetCurrentDirectoryA Path
  Else
    If Mid(Path, 2, 1) = ":" Then ChDrive Left(Path, 2)
    ChDir Path
  End If
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
  Shell """" & GitExeFolder & "\cmd\Gitk.exe"" --all"
  
  GoToNameColumn
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
  Shell """" & GitExeFolder & "\Git-bash.exe"""
  
  GoToNameColumn
End Sub

Function FolderName(FullPath As String) As String
  FolderName = Mid(FullPath, InStrRev(FullPath, "\") + 1)
End Function

Function GetFilesIn(Folder As String) As Collection
  Dim F As String
  Set GetFilesIn = New Collection
  F = Dir(Folder & "\*")
  Do While F <> ""
    GetFilesIn.Add F
    F = Dir
  Loop
End Function

Function GetFoldersIn(Folder As String) As Collection
  Dim F As String
  Set GetFoldersIn = New Collection
  F = Dir(Folder & "\*", vbDirectory)
  Do While F <> ""
    If GetAttr(Folder & "\" & F) And vbDirectory Then GetFoldersIn.Add F
    F = Dir
  Loop
End Function

Sub TestQuickSortArrayKV()
  Dim I As Integer, Arr(3 To 12) As KeyValue, N As Integer
  For I = LBound(Arr) To UBound(Arr)
    N = Rnd * UBound(Arr)
    Set Arr(I) = NewKeyValue(N, N)
  Next I
  
  QuickSortArrayKV Arr
  
  PrintArray Arr
End Sub

Sub QuickSortArrayKVGetValues(Arr() As KeyValue, Result)
  QuickSortArrayKV Arr
  
  Dim I As Integer
  ReDim Result(UBound(Arr))
  For I = LBound(Arr) To UBound(Arr)
    Set Result(I) = Arr(I).Value
  Next I
End Sub

Sub QuickSortArrayKV(Arr() As KeyValue, Optional IStart As Integer = -999, Optional IEnd As Integer)
  If IStart = -999 Then
    IStart = LBound(Arr)
    IEnd = UBound(Arr)
  End If
  
  Dim I As Integer, K As Integer, PivotKey
  I = IStart
  K = IEnd
  
  If IEnd - IStart >= 1 Then
    PivotKey = Arr(IStart).Key
    
    Do While K > I
      Do While Arr(I).Key <= PivotKey And I <= IEnd And K > I
        I = I + 1
      Loop
  
      Do While Arr(K).Key > PivotKey And K >= IStart And K >= I
        K = K - 1
      Loop
      
      If K > I Then SwapArrayKV Arr, I, K
    Loop
    
    SwapArrayKV Arr, IStart, K
    
    QuickSortArrayKV Arr, IStart, K - 1
    QuickSortArrayKV Arr, K + 1, IEnd
  End If
End Sub

Sub QuickSortArray(Arr(), Optional IStart As Integer = -999, Optional IEnd As Integer)
  If IStart = -999 Then
    IStart = LBound(Arr)
    IEnd = UBound(Arr)
  End If
  
  Dim I As Integer, K As Integer, PivotKey
  I = IStart
  K = IEnd
  
  If IEnd - IStart >= 1 Then
    PivotKey = Arr(IStart)
    
    Do While K > I
      Do While Arr(I) <= PivotKey And I <= IEnd And K > I
        I = I + 1
      Loop
  
      Do While Arr(K) > PivotKey And K >= IStart And K >= I
        K = K - 1
      Loop
      
      If K > I Then SwapArray Arr, I, K
    Loop
    
    SwapArray Arr, IStart, K
    
    QuickSortArray Arr, IStart, K - 1
    QuickSortArray Arr, K + 1, IEnd
  End If
End Sub

Function QuickSort(ByVal Coll As Collection) As Collection
  If Coll.Count <= 1 Then
    Set QuickSort = Coll
    Exit Function
  End If
  
  Dim Smaller As New Collection, Bigger As New Collection
  Dim Pivot As Variant, N As Long, V As Variant
  
  N = Coll.Count / 2
  Pivot = Coll(N)
  Coll.Remove N
  
  Do While Coll.Count
    If Coll(1) < Pivot Then Smaller.Add Coll(1) Else Bigger.Add Coll(1)
    Coll.Remove 1
  Loop
  
  Set QuickSort = New Collection
  
  For Each V In QuickSort(Smaller)
    QuickSort.Add V
  Next V
  
  QuickSort.Add Pivot

  For Each V In QuickSort(Bigger)
    QuickSort.Add V
  Next V
End Function

Function QuickSortKVGetValues(ByVal Coll As Collection) As Collection
  Dim C1 As Collection, C2 As New Collection, V As Variant
  Set C1 = QuickSortKV(Coll)
  For Each V In C1
    C2.Add V.Value
  Next V
  Set QuickSortKVGetValues = C2
End Function

Function QuickSortKV(ByVal Coll As Collection) As Collection
  If Coll.Count <= 1 Then
    Set QuickSortKV = Coll
    Exit Function
  End If
  
  Dim Smaller As New Collection, Bigger As New Collection
  Dim Pivot As KeyValue, N As Long, V As KeyValue
  
  N = Coll.Count / 2
  Set Pivot = Coll(N)
  Coll.Remove N
  
  Do While Coll.Count
    If Coll(1).Key < Pivot.Key Then Smaller.Add Coll(1) Else Bigger.Add Coll(1)
    Coll.Remove 1
  Loop
  
  Set QuickSortKV = New Collection
  
  For Each V In QuickSortKV(Smaller)
    QuickSortKV.Add V
  Next V
  
  QuickSortKV.Add Pivot

  For Each V In QuickSortKV(Bigger)
    QuickSortKV.Add V
  Next V
End Function

Function NewKeyValue(Key As Variant, Value As Variant) As KeyValue
  Set NewKeyValue = New KeyValue
  NewKeyValue.Init Key, Value
End Function

Sub SwapArrayKV(Arr() As KeyValue, I1 As Integer, I2 As Integer)
  Dim O As Object
  Set O = Arr(I1)
  Set Arr(I1) = Arr(I2)
  Set Arr(I2) = O
End Sub

Sub SwapArray(Arr(), I1 As Integer, I2 As Integer)
  Dim X
  X = Arr(I1)
  Arr(I1) = Arr(I2)
  Arr(I2) = X
End Sub

Sub PrintArray(Arr() As KeyValue)
  Dim I As Integer
  For I = LBound(Arr) To UBound(Arr)
    Debug.Print Arr(I).Value;
  Next I
  Debug.Print
End Sub
