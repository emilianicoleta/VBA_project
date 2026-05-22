Attribute VB_Name = "modRaportErori"
Sub InitRaportErori()
    Dim ws As Worksheet
    On Error Resume Next
    Application.DisplayAlerts = False
    Worksheets("Raport_Erori").Delete
    Application.DisplayAlerts = True
    On Error GoTo 0
    
    Set ws = Worksheets.Add
    ws.Name = "Raport_Erori"
    
    ws.Range("A1:E1").Value = Array("Fisier", "Rand", "Camp", "Tip_Eroare", "Detalii")
    ws.Rows(1).Font.Bold = True
End Sub

Sub LogEroare(numeFisier As String, rand As Long, camp As String, _
              tipEroare As String, detalii As String)
    Dim ws As Worksheet
    Dim nextRow As Long
    
    Set ws = Worksheets("Raport_Erori")
    nextRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row + 1
    
    ws.Cells(nextRow, 1).Value = numeFisier
    ws.Cells(nextRow, 2).Value = rand
    ws.Cells(nextRow, 3).Value = camp
    ws.Cells(nextRow, 4).Value = tipEroare
    ws.Cells(nextRow, 5).Value = detalii
End Sub


