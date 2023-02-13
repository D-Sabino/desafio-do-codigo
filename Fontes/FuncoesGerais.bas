Attribute VB_Name = "FuncoesGerais"
'Codigo encontrado na internet!
Public Sub PadronizarCrystalReport(frmFormulario As Form, strNomeControle As String)
    On Error GoTo TrataErro
    
    Dim intContador As Integer
    
    For intContador = 0 To frmFormulario.Controls.Count - 1
        If TypeOf frmFormulario.Controls(intContador) Is CrystalReport And _
            UCase$(frmFormulario.Controls(intContador).Name) = UCase$(strNomeControle) Then
            Exit For
        End If
    Next intContador

    If intContador <> frmFormulario.Controls.Count - 1 Then
        frmFormulario.Controls(intContador).WindowShowCloseBtn = True
        frmFormulario.Controls(intContador).WindowShowExportBtn = True
        frmFormulario.Controls(intContador).WindowShowGroupTree = False
        frmFormulario.Controls(intContador).WindowShowNavigationCtls = True
        frmFormulario.Controls(intContador).WindowShowPrintBtn = True
        frmFormulario.Controls(intContador).WindowShowPrintSetupBtn = True
        frmFormulario.Controls(intContador).WindowShowRefreshBtn = True
        frmFormulario.Controls(intContador).WindowShowSearchBtn = True
        frmFormulario.Controls(intContador).WindowShowZoomCtl = True
        frmFormulario.Controls(intContador).WindowShowProgressCtls = True
    End If

    Exit Sub
    Resume
TrataErro:
    MsgBox "Ocorreu um erro ao PadronizarCrystalReport. PadronizarCrystalReport, FuncoesGerais"
End Sub

