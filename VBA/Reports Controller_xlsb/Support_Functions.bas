Attribute VB_Name = "Support_Functions"
Option Explicit
Option Compare Text

Function URLEncodeString(str As String) As String
    If Val(Application.Version) >= 15 Then
        URLEncodeString = WorksheetFunction.EncodeURL(str)
    Else
        ' EncodeURL is not available in prev versions
        URLEncodeString = Support_Functions.URLEncode(str)
    End If
End Function

' http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
Function URLEncode( _
   StringVal As String, _
   Optional SpaceAsPlus As Boolean = False) As String

  Dim StringLen As Long: StringLen = Len(StringVal)

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

Function IsEditing() As Boolean
' checks if Excel is in Edit cell mode
    On Error GoTo ErrHandler
    If Application.Interactive = True Then
        On Error Resume Next
        Application.Interactive = False
        Application.Interactive = True
        If Err.Number <> 0 Then
            IsEditing = True
            Err.Clear
        End If
    Else
        ' false
    End If
    
    Exit Function
    
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now, "IsEditing", Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Function

Function decodeURL(str As String) As String
    
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

