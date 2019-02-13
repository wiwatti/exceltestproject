VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub CommandButton1_Click()

Dim i As Integer
i = 1

Do While i < 6
    Cells(i, 1).Value = 20
    i = i + 1
Loop

i = 1

Do While Cells(i, 1) <> ""
    Cells(i, 2).Value = Cells(i, 1).Value + 10
    i = i + 2
Loop



End Sub
