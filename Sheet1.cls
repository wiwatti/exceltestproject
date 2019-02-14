VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()

Dim i As Integer

For i = 1 To 6
    Cells(i, 1).Value = 100
    Cells(i, 1).Value = 200
    Cells(i, 1).Value = 300
Next i

End Sub

Private Sub CommandButton2_Click()

End Sub
