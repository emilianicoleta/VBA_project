Attribute VB_Name = "modValidari"
Option Explicit

Sub ImportaSiValideazaFisiere(folderPath As String)
    Dim f As String
    Dim wbSrc As Workbook
    Dim wsSrc As Worksheet
    Dim total As Long, count As Long

    ' 1. Numaram fi?ierele din folder
    f = Dir(folderPath & "*.xlsx")
    Do While f <> ""
        total = total + 1
        f = Dir
    Loop

    If total = 0 Then
        MsgBox "Nu exista fi?iere .xlsx în folderul selectat.", vbExclamation
        Exit Sub
    End If

    ' 2. Resetam progress bar-ul
    Call UpdateProgress(0)

    ' 3. Ini?ializam raportul de erori
    Call InitRaportErori

    ' 4. Parcurgem fi?ierele ?i actualizam progress bar-ul
    count = 0
    f = Dir(folderPath & "*.xlsx")

    Do While f <> ""
        count = count + 1
        Call UpdateProgress(count / total)

        Set wbSrc = Workbooks.Open(folderPath & f)
        Set wsSrc = wbSrc.Sheets(1)

        Call ValideazaFoaie(wsSrc, f)

        wbSrc.Close SaveChanges:=False
        f = Dir
    Loop

    ' 5. Progress bar complet
    Call UpdateProgress(1)

End Sub


Sub ValideazaFoaie(ws As Worksheet, numeFisier As String)
    Dim lastRow As Long, i As Long
    Dim soldInit As Double, intrari As Double, iesiri As Double, soldFinal As Double

    lastRow = ws.Cells(ws.Rows.count, "A").End(xlUp).Row

    ' 1. Verificari de completitudine
    For i = 2 To lastRow
        If IsEmpty(ws.Cells(i, "A")) Then
            Call LogEroare(numeFisier, i, "A", "Lipsa", "ID lipsa")
        End If
        If IsEmpty(ws.Cells(i, "B")) Or IsEmpty(ws.Cells(i, "E")) Then
            Call LogEroare(numeFisier, i, "B/E", "Lipsa", "Sold ini?ial sau final lipsa")
        End If
    Next i

    ' 2. Verificari logice
    For i = 2 To lastRow
        If Not IsEmpty(ws.Cells(i, "B")) And Not IsEmpty(ws.Cells(i, "C")) _
           And Not IsEmpty(ws.Cells(i, "D")) And Not IsEmpty(ws.Cells(i, "E")) Then

           soldInit = ws.Cells(i, "B").Value
           intrari = ws.Cells(i, "C").Value
           iesiri = ws.Cells(i, "D").Value
           soldFinal = ws.Cells(i, "E").Value

           If Abs(soldInit + intrari - iesiri - soldFinal) > 0.01 Then
               Call LogEroare(numeFisier, i, "B:E", "Incoerenta", _
                   "SoldFinal diferit de SoldInitial + Intrari - Iesiri")
               ws.Rows(i).Interior.Color = vbYellow
           End If
        End If
    Next i

    ' 3. Detectare outlieri pe coloana C
    Call DetecteazaOutlieri(ws, numeFisier, "C", 2, lastRow)

End Sub


