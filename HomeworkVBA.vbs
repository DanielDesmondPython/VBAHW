Attribute VB_Name = "Module2"
Sub lasttry()
Dim ws As Worksheet
Dim Ticker As String
Dim totalvol As Double
totalvol = 0
Dim outputrow As Integer

Dim openstock As Double
Dim closestock As Double
Dim yearly_change As Double
Dim lastrow As Long

lastrow = ActiveSheet.Cells(Rows.Count, 1).End(xlUp).Row
Ticker = Cells(2, 1).Value
Dim percent_change As Double

For Each ws In Sheets

openstock = Cells(2, 3).Value
closestock = Cells(2, 5).Value
Cells(1, 9).Value = ÒTickerÓ
Cells(1, 10).Value = Òyearly_ChangeÓ
Cells(1, 11).Value = Òpercent_changeÓ
Cells(1, 12).Value = ÒtotalvolÓ
outputrow = 2
For i = 2 To lastrow
Cells(i, 11).NumberFormat = "0.00"

If Cells(i, 10).Value > 0 Then
Cells(i, 10).Interior.ColorIndex = 4
ElseIf Cells(i, 10).Value < 0 Then
Cells(i, 10).Interior.ColorIndex = 3
End If

If (openstock > 0) And (closestock > 0) Then
Cells(outputrow, 11).Value = Cells(outputrow, 10).Value / openstock
End If
If Cells(i, 1).Value = Cells(i + 1, 1).Value Then
totalvol = totalvol + Cells(i, 7).Value
Else
Cells(outputrow, 9).Value = Cells(i - 1, 1).Value
closestock = Cells(i - 1, 6).Value
openstock = Cells(i + 1, 3).Value
Cells(outputrow, 10).Value = openstock - closestock
totalvol = Cells(i, 7).Value + totalvol
Cells(outputrow, 12).Value = totalvol
outputrow = outputrow + 1
totalvol = 0
End If

Next i

Next ws

End Sub


End Sub

