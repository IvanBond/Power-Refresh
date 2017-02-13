Attribute VB_Name = "Support_Functions"
Option Explicit

' http://www.fmsinc.com/microsoftaccess/modules/examples/avoiddoevents.asp
' http://analystcave.com/vba-sleep-vs-wait/
' http://www.exceltrick.com/formulas_macros/vba-wait-and-sleep-functions/
Public Sub WaitSeconds(intSeconds As Integer)
    Dim datTime As Date

    datTime = DateAdd("s", intSeconds, Now)

    Do
        ' Yield to other programs (better than using DoEvents which eats up all the CPU cycles)
        Sleep 100
        DoEvents
    Loop Until Now >= datTime
End Sub

Public Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False) As String

  Dim StringLen As Long
  StringLen = Len(StringVal)

  If StringLen > 0 Then
    ReDim result(StringLen) As String
    Dim i As Long, CharCode As Integer
    Dim Char As String, Space As String

    If SpaceAsPlus Then Space = "+" Else Space = "%20"

    For i = 1 To StringLen
      Char = Mid$(StringVal, i, 1)
      CharCode = Asc(Char)
      Select Case CharCode
        Case 97 To 122, 65 To 90, 48 To 57, 45, 46, 95, 126
          result(i) = Char
        Case 32
          result(i) = Space
        Case 0 To 15
          result(i) = "%0" & Hex(CharCode)
        Case Else
          result(i) = "%" & Hex(CharCode)
      End Select
    Next i
    URLEncode = Join(result, "")
  End If
End Function

Public Function decodeURL(str As String) As String
    
     ' =SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(SUBSTITUTE(CELL_WITH_URL,”%3F”,”?”),”%20?,” “),”%25”, “%”),”%26?,”&”),”%3D”,”=”),”%7B”,”{“),”%7D”,”}”),”%5B”,”[“),”%5D”,”]”)
     
    ' https://excelsnippets.wordpress.com/2011/03/29/excel-vba-function-to-revert-url-encoded-strings-to-regular-strings/
    Dim i As Integer
    Dim txt As String
    Dim hexChr As String

   txt = str

   'replace '+' with space
   txt = Replace(txt, "+", " ")

   For i = 1 To 255
      Select Case i
         Case 1 To 15
            hexChr = "%0" & Hex(i)
         Case 37
            'skip '%' character
         Case Else
            hexChr = "%" & Hex(i)
      End Select

      txt = Replace(txt, UCase(hexChr), Chr(i))
      txt = Replace(txt, LCase(hexChr), Chr(i))
   Next

   'replace '%' character
   txt = Replace(txt, "%25", "%")
   decodeURL = txt
   
End Function


