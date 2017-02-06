Attribute VB_Name = "Globals"
Option Explicit

Public Control_Table As ListObject
Public LOG_Table As ListObject
Public Excel_Path As String
Public Refresher_Path As String
'

Sub Set_Global_Variables()
    
    Set Control_Table = ThisWorkbook.Sheets("ControlPanel").ListObjects("ControlTable")
    Set LOG_Table = ThisWorkbook.Sheets("LOG").ListObjects("LOG_Table")
        
    Excel_Path = Application.path & "\EXCEL.EXE"
    ' "C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE"
    
    Refresher_Path = ThisWorkbook.path & "\Refresher.xlsb"
    
End Sub
