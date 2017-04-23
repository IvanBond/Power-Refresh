Attribute VB_Name = "Support_Funcitons"
Option Explicit

' http://stackoverflow.com/questions/218181/how-can-i-url-encode-a-string-in-excel-vba
Public Function URLEncode( _
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
ErrHandler:
    If Err.Number <> 0 Then
        Debug.Print Now() & ": IsEditing: " & Err.Number & ": " & Err.Description
        Err.Clear
    End If
End Function
