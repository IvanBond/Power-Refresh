Attribute VB_Name = "Globals"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#Else
    Public Declare Sub Sleep Lib "kernel32" (ByVal lngMilliSeconds As Long)
#End If

Public Control_Table As ListObject
Public LOG_Table As ListObject
Public Excel_Path As String
Public Refresher_Path As String
'

Sub Set_Global_Variables()
    
    Set Control_Table = ControlPanel.ListObjects("ControlTable")
    Set LOG_Table = Logs.ListObjects("LOG_Table")
    
    Excel_Path = Application.path & "\EXCEL.EXE"
    ' "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    
    Refresher_Path = ThisWorkbook.path & "\Refresher.xlsb"
    
End Sub
