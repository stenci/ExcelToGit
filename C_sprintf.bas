Attribute VB_Name = "C_sprintf"
Option Explicit

' #VBIDEUtils#************************************************************
' * Programmer Name  : Thierry Waty
' * Web Site         : http://www.vbdiamond.com
' * E-Mail           : waty.thierry@vbdiamond.com
' * Date             : 01/10/2001
' * Time             : 13:38
' **********************************************************************
' * Comments         : Simulate in VB the "sprintf" function in C (updated)
' *
' * Simulate in VB the "sprintf" function in C
' *
' **********************************************************************
Const NONE = 0
Const STRINGTYPE = 1
Const INTEGERTYPE = 2
Const LONGTYPE = 3
Const FLOATTYPE = 4
Const CHARPERCENT = 5
Const HEXTYPE = 6

Function SPrintf2(Mask As String, ParamArray Tokens()) As String
  'SPrintf2("ab {1} de {0} fg", "XX", 123)
  'Result:  "ab 123 de XX fg"
  Dim I As Integer
  SPrintf2 = Mask
  For I = 0 To UBound(Tokens)
    SPrintf2 = Replace(SPrintf2, "{}", "{" & I & "}", , 1)
  Next
  For I = 0 To UBound(Tokens)
    SPrintf2 = Replace(SPrintf2, "{" & I & "}", Tokens(I))
  Next
End Function

Public Function SPrintf(sFormats As String, ParamArray aArguments() As Variant) As String

   Dim nCurrentFlag     As Long
   Dim nPos             As Long
   Dim sWork            As String
   Dim nCurVal          As Long
   Dim nMaxArg          As Integer
   Dim sCurFormat       As String
   Dim nArgCount        As Integer
   Dim nxIndex          As Long
   Dim bFound           As Boolean
   Dim nType            As Integer
   Dim sCurParm         As String
   Dim nLenFlags        As Long
   Dim bUpper           As Boolean

   ' If an array is passed, replace the ParamArray with it
   If UBound(aArguments) >= LBound(aArguments) Then
      If IsArray(aArguments(0)) Then
         aArguments = aArguments(0)
      End If
   End If

   ' *** Get the number of arguments
   nMaxArg = UBound(aArguments) + 1

   ' *** Length of the flags
   nLenFlags = Len(sFormats)

   ' *** Initialize some variables
   nCurrentFlag = 1
   nCurVal = 0
   nArgCount = 0

   ' *** Get the first flag
   nPos = InStr(nCurrentFlag, sFormats, "%")

   ' *** Verify if the number of flags is the same as the number of argument
   Do While (nPos > 0)
      If Mid$(sFormats, nPos + 1, 1) <> "%" Then ' *** Don't count %%, will be converted to % later
         nArgCount = nArgCount + 1
         nCurrentFlag = nPos + 1
      Else
         nCurrentFlag = nPos + 2
      End If

      ' *** Get next flag
      nPos = InStr(nCurrentFlag, sFormats, "%")
   Loop

   ' *** Compare the number of flags against the number of arguments
   If nArgCount <> nMaxArg Then Err.Raise 450, , "Mismatch of parameters for string " & sFormats & ".  Expected " & nArgCount & " but received " & nMaxArg & "."

   ' *** Initialize some variables
   nCurrentFlag = 1
   nCurVal = 0
   nArgCount = 0
   sWork = ""

   ' *** Get the first flag
   nPos = InStr(nCurrentFlag, sFormats, "%")

   Do While (nPos > 0)
      ' *** First, get the variable identifier.
      ' *** Scan from nCurrentFlag (the %) to EOL looking for the
      ' ***  first occurance of s, d, l, or f
      bFound = False
      nType = NONE
      nxIndex = nPos + 1
      Do While (bFound = False) And (nxIndex <= nLenFlags)
         If Not bFound Then
            sCurParm = Mid$(sFormats, nxIndex, 1)
            Select Case Mid$(sFormats, nxIndex, 1)
               Case "%"
                  nType = CHARPERCENT
                  bUpper = False
                  bFound = True
                  nPos = nPos + 1
                  nCurVal = nxIndex + 2
               Case "s"
                  nType = STRINGTYPE
                  bUpper = False
                  bFound = True
                  nCurVal = nxIndex + 1
               Case "S"
                  nType = STRINGTYPE
                  bUpper = True
                  bFound = True
                  nCurVal = nxIndex + 1
               Case "d", "i", "u"
                  nType = INTEGERTYPE
                  bUpper = False
                  bFound = True
                  nCurVal = nxIndex + 1
               Case "l"
                  If Mid$(sFormats, nxIndex + 1, 1) = "d" Then
                     nType = LONGTYPE
                     bUpper = False
                     bFound = True
                     nCurVal = nxIndex + 2
                  Else
                     Err.Raise 93, , "Unrecognized pattern " & Mid$(sFormats, nxIndex - 1, 3) & " in " & sFormats
                  End If
               Case "f", "e", "g"
                  nType = FLOATTYPE
                  bUpper = False
                  bFound = True
                  nCurVal = nxIndex + 1
               Case "E", "G"
                  nType = FLOATTYPE
                  bUpper = False
                  bFound = True
                  nCurVal = nxIndex + 1
               Case "x"
                  nType = HEXTYPE
                  bUpper = False
                  bFound = True
                  nCurVal = nxIndex + 1
               Case "X"
                  nType = HEXTYPE
                  bUpper = True
                  bFound = True
                  nCurVal = nxIndex + 1
            End Select
         End If

         If Not bFound Then nxIndex = nxIndex + 1

      Loop

      ' *** Not found, raise an error
      If Not bFound Then Err.Raise 93, , "Invalid % format in <" & sFormats & ">"

      ' *** Get the complete flag
      sCurParm = Mid$(sFormats, nPos, nCurVal - nPos)

      ' *** Different case if Percent or other
      If nType = CHARPERCENT Then
         sWork = sWork & Mid$(sFormats, nCurrentFlag, nPos - nCurrentFlag)
         nCurVal = nCurVal - 1
      Else
         sCurFormat = BuildFormat(sCurParm, nType, bUpper, aArguments(nArgCount))
         sWork = sWork & Mid$(sFormats, nCurrentFlag, nPos - nCurrentFlag) & sCurFormat
         nArgCount = nArgCount + 1
      End If
      nCurrentFlag = nCurVal

      ' *** Get next flag
      nPos = InStr(nCurrentFlag, sFormats, "%")
   Loop

   ' *** If there are no more flags, add the rest of the string and get out
   sWork = sWork & Mid$(sFormats, nCurrentFlag)

   SPrintf = TreatBackSlash(sWork)

End Function

Function BuildFormat(sFormat As String, nDataType As Integer, bUpperCase As Boolean, vCurrentValue As Variant) As String
   ' *** Build the format

   Dim sPrefix          As String
   Dim sFlag            As String
   Dim nWidth           As Long
   Dim bAlignLeft       As Boolean
   Dim bSign            As Boolean
   Dim sPad             As String * 1
   Dim bBlank           As Boolean
   Dim bDecimal         As Boolean
   Dim nI               As Integer
   Dim sTmp             As String
   Dim sWidth           As String
   Dim nPrecision       As Integer
   Dim nPlaces          As Integer
   Dim NUnits           As Integer
   Dim sCurrentValue    As Variant

   If (Len(sFormat) < 2) Then
      BuildFormat = ""
      Exit Function
   End If

   ' *** Get the flag
   sFlag = ""
   bAlignLeft = False
   bSign = False
   sPad = "@"
   bBlank = False
   bDecimal = False
   Select Case Mid$(sFormat, 2, 1)
      Case "-":
         bAlignLeft = True
         sFormat = Mid$(sFormat, 3)

      Case "+":
         bSign = True
         sFormat = Mid$(sFormat, 3)

      Case "0":
         sPad = "0"
         sFormat = Mid$(sFormat, 3)

      Case " ":
         bBlank = True
         sFormat = Mid$(sFormat, 3)

      Case "#":
         bDecimal = True
         sFormat = Mid$(sFormat, 3)

      Case Else
         sFormat = Mid$(sFormat, 2)

   End Select

   ' *** Get the width
   If nDataType = LONGTYPE Then
      sPrefix = Mid$(sFormat, 1, Len(sFormat) - 2)
   Else
      sPrefix = Mid$(sFormat, 1, Len(sFormat) - 1)
   End If

   ' *** Get the width
   sWidth = ""
   nI = 1
   sTmp = Mid$(sPrefix, nI, 1)
   Do While IsNumeric(sTmp)
      sWidth = sWidth & sTmp

      nI = nI + 1
      sTmp = Mid$(sPrefix, nI, 1)
   Loop

   If (Trim$(sWidth) = "") Then sWidth = "0"
   nWidth = CLng(sWidth)

   ' *** Check the precision
   nPrecision = InStr(sPrefix, ".")
   If (nPrecision = 0) Then
      ' *** No precision, only width (eventually)
      If (bAlignLeft = False) Then
         sFormat = String(nWidth, sPad)
      Else
         If (Len(CStr(vCurrentValue)) > nWidth) Then nWidth = Len(CStr(vCurrentValue))
         sFormat = String(Len(CStr(vCurrentValue)), sPad) & String(nWidth - Len(CStr(vCurrentValue)), " ")
      End If
   Else
      sTmp = "0"
      nI = nPrecision + 1
      Do While IsNumeric(Mid$(sPrefix, nI, 1))
         sTmp = sTmp & Mid$(sPrefix, nI, 1)
         nI = nI + 1
      Loop

      nPlaces = CLng(sTmp)
      
      If nWidth < nPlaces Then
        If vCurrentValue Then
          NUnits = Int(Log(Abs(vCurrentValue)) / Log(10))
          If NUnits < 0 Then NUnits = 0
          nWidth = nPlaces + 1 + NUnits + 1 + IIf(vCurrentValue < 0, 1, 0)
        Else
          nWidth = nPlaces + 2
        End If
      End If

      Select Case nDataType
         Case INTEGERTYPE, LONGTYPE, HEXTYPE:
            ' *** Take the right 'nWidth' characters because format with insert at least one space
            sFormat = Right$(Format$(" ", String$(nWidth - nPlaces, sPad)) & String$(nPlaces, "0"), nWidth)
         Case FLOATTYPE:
            sFormat = String$(nWidth - nPlaces - 2, "#") & "0." & String$(nPlaces, "0")
      End Select

   End If

   If nDataType = HEXTYPE Then
      ' *** Convert to Hex
      sCurrentValue = Hex$(vCurrentValue)

      ' *** Display the entire number even if the format is smaller
      If Len(sFormat) < Len(sCurrentValue) Then
         sFormat = vbNullString
         ' *** Else set the current value equal to the 0 padded string (if it's not 0 padded,
         ' *** the format works correctly already)
      ElseIf nPrecision <> 0 Or sPad = "0" Then
         sCurrentValue = Left$(sFormat, Len(sFormat) - Len(sCurrentValue)) & sCurrentValue
         sFormat = vbNullString
      End If

   Else
      sCurrentValue = vCurrentValue
   End If

   If nDataType <> STRINGTYPE Then
      If bUpperCase Then
         sCurrentValue = UCase(sCurrentValue)
      Else
         sCurrentValue = LCase(sCurrentValue)
      End If
   End If

   If sFormat = vbNullString Then
      BuildFormat = sCurrentValue
   Else
      BuildFormat = Format$(sCurrentValue, sFormat)
      If (nWidth - Len(BuildFormat)) < 0 Then
         BuildFormat = String(nWidth, "#")
      Else
         BuildFormat = String((nWidth - Len(BuildFormat)), " ") & BuildFormat
      End If
   End If

End Function

Public Function TreatBackSlash(sLine As String) As String
   ' *** Treat all the backslach

   Dim nPos             As Long
   Dim sChar            As String * 1

   ' *** Get the first backslash
   nPos = InStr(sLine, "\")

   Do While (nPos > 0)
      ' *** First, get the char after
      sChar = Mid$(sLine, nPos + 1, 1)
      Select Case sChar
         Case "n"
            sLine = Left$(sLine, nPos - 1) & Chr$(13) & Chr$(10) & Right$(sLine, Len(sLine) - nPos - 1)
            nPos = nPos + 1
         Case "r"
            sLine = Left$(sLine, nPos - 1) & Chr$(13) & Right$(sLine, Len(sLine) - nPos - 1)
            nPos = nPos + 1
         Case "t"
            sLine = Left$(sLine, nPos - 1) & Chr$(9) & Right$(sLine, Len(sLine) - nPos - 1)
            nPos = nPos + 1
         Case "\"
            sLine = Left$(sLine, nPos - 1) & "\" & Right$(sLine, Len(sLine) - nPos - 1)
            nPos = nPos + 1
         Case Else
            ' If there is not a recognizable flag, then take out the slash
            sLine = Left$(sLine, nPos - 1) & Right$(sLine, Len(sLine) - nPos)
            nPos = nPos + 1
      End Select

      nPos = InStr(nPos, sLine, "\")
   Loop

   TreatBackSlash = sLine

End Function
