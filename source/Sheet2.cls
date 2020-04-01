VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub MergeCells()
    Dim i As Byte
    Dim A As String
    Dim B As String
    
    For i = 2 To 160 Step 2
        A = "B" + Trim(Str(i)) + ":B" + Trim(Str(i + 1))
        B = "K" + Trim(Str(i)) + ":K" + Trim(Str(i + 1))
        Sheet2.Range(A + "," + B).Merge
    Next
    
End Sub


Private Sub Chart_Clear()
    If Sheet2.ChartObjects.Count > 0 Then
        Sheet2.ChartObjects.Delete
    End If
End Sub

