VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)
  Dim Txt As String, C As Integer
  C = Target.Column
  If C = COL_EXPORT Or C = COL_GIT_GUI Or C = COL_GITK Or C = COL_GIT_BASH Then
    Txt = Target.Cells(1, 1).Value2
    Select Case Txt
      Case "Refresh list of open documents": Refresh
      Case "Export": Export
      Case "Git gui": GitGui
      Case "gitk": Gitk
      Case "bash": GitBash
    End Select
  ElseIf C = COL_GIT_FOLDER Or C = COL_FOLDER Then
    Txt = Cells(Target.Row, C).Value2
    OpenFolder Txt
  End If
End Sub
