Attribute VB_Name = "Module19"
Sub findDepWithNoCoverageMHP()

Dim ws As Worksheet
Dim lr As Integer, lc As Integer, i As Integer, j As Integer

Set ws = Sheets("MHPPartDepAndCov")
lr = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lr
    Select Case ws.Cells(i, 16).Value
        Case "C"
            For j = 21 To 33
                If ws.Cells(i, j).Value = "SDP" Then
                    ws.Cells(i, 47).Value = "Dependent Life exists."
                    Exit For
                End If
            Next j
        Case "S"
            For j = 21 To 33
                If ws.Cells(i, j).Value = "SSP" Then
                    ws.Cells(i, 47).Value = "Spouse Life exists."
                    Exit For
                End If
            Next j
        Case Else
            ws.Cells(i, 47).Value = "Check Dependent Relationship."
    End Select
Next i

For i = 2 To lr
    If ws.Cells(i, 47).Value = "" Then
        lc = ws.Cells(i, Columns.Count).End(xlToLeft).Column
        For j = 34 To lc
            If ws.Cells(i, j).Value <> "P00" Then
                ws.Cells(i, 47).Value = "Non-Employee Coverage exists."
                Exit For
            End If
        Next j
    End If
Next i

MsgBox "Done"

End Sub

