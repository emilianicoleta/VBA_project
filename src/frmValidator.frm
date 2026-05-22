VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmValidator 
   Caption         =   "UserForm1"
   ClientHeight    =   4755
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9390.001
   OleObjectBlob   =   "frmValidator.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmValidator"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdBrowse_Click()
    Dim fldr As FileDialog
    Set fldr = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fldr
        .Title = "Selecteaza folderul cu fi?iere"
        .AllowMultiSelect = False
        If .Show = -1 Then
            txtFolderPath.Text = .SelectedItems(1) & "\"
        End If
    End With
End Sub
Private Sub cmdClose_Click()
    Unload Me
End Sub
Private Sub cmdRun_Click()
    If txtFolderPath.Text = "" Then
        MsgBox "Selecteaza un folder înainte de a rula validarea.", vbExclamation
        Exit Sub
    End If
    
    Call ImportaSiValideazaFisiere(txtFolderPath.Text)
    
    MsgBox "Validare finalizata. Verifica foaia 'Raport_Erori'.", vbInformation
End Sub

