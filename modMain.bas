Attribute VB_Name = "modMain"
Public Sub Main()

    If Right(LCase(Command()), 4) = ".lst" Then
        Load frmGermatik
        frmGermatik.Show
'        frmGermatik.CommonDialogList.fileName = GetLongFilename(Command())
        frmGermatik.CommonDialogList.fileName = Command()
        Call frmGermatik.doOpen
        Call frmGermatik.mnuStart_Click
    Else
        Load frmGermatik
        frmGermatik.Hide
        frmSplash.Show vbModal
        frmGermatik.Show
    End If

End Sub
