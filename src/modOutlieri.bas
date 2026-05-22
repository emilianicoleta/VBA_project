Attribute VB_Name = "modOutlieri"
Sub DetecteazaOutlieri(ws As Worksheet, numeFisier As String, _
                       col As String, firstRow As Long, lastRow As Long)
    Dim rng As Range, cell As Range
    Dim medie As Double, devStd As Double
    
    Set rng = ws.Range(col & firstRow & ":" & col & lastRow)
    
    medie = WorksheetFunction.Average(rng)
    devStd = WorksheetFunction.StDev(rng)
    
    For Each cell In rng
        If Not IsEmpty(cell) Then
            If devStd > 0 Then
                If Abs(cell.Value - medie) > 3 * devStd Then
                    Call LogEroare(numeFisier, cell.Row, col, "Outlier", _
                        "Valoare atipica fa?a de distribu?ia coloanei " & col)
                    cell.Interior.Color = vbRed
                End If
            End If
        End If
    Next cell
End Sub


