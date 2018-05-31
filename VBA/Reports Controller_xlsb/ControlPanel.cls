VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ControlPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Sub ResetComments()
    Dim pComment As Comment
    For Each pComment In Me.Comments
        pComment.Shape.Top = pComment.Parent.Top + 5
        pComment.Shape.Left = pComment.Parent.offset(0, 1).Left + 5
    Next
End Sub

