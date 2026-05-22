Attribute VB_Name = "modMain"
Sub StartValidator()
    frmValidator.Show
End Sub

Sub UpdateProgress(percent As Double)
    Dim targetWidth As Double
    Dim currentWidth As Double
    Dim stepSize As Double
    
    ' Bara ?inta (în func?ie de procent)
    targetWidth = frmValidator.fraProgress.Width * percent
    
    ' Bara actuala
    currentWidth = frmValidator.lblProgress.Width
    
    ' Pasul de anima?ie (smooth)
    stepSize = (targetWidth - currentWidth) / 10
    
    ' Anima?ie în 10 pa?i mici
    Dim i As Integer
    For i = 1 To 10
        frmValidator.lblProgress.Width = frmValidator.lblProgress.Width + stepSize
        frmValidator.lblPercent.Caption = Format(percent * 100, "0") & "%"
        DoEvents
    Next i
End Sub



