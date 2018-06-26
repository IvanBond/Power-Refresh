Attribute VB_Name = "z_API"
Option Explicit

#If VBA7 Then
    Public Declare PtrSafe Function ShowWindow Lib "user32.dll" _
        (ByVal HWND As LongPtr, ByVal nCmdShow As LongPtr) As LongPtr
#Else
    Public Declare Function ShowWindow Lib "user32.dll" _
        (ByVal HWND As Long, ByVal nCmdShow As Long) As Long
#End If

Public Const SW_HIDE As Long = 0
Public Const SW_SHOW As Long = 5
