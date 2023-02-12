VERSION 5.00
Begin VB.Form frmPrincipal 
   Caption         =   "Principal"
   ClientHeight    =   7395
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   12945
   LinkTopic       =   "Form1"
   ScaleHeight     =   7395
   ScaleWidth      =   12945
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
    On Error GoTo TrataErro
    
    'Inicializando a conexão com o servidor.
    Dim cn As New ADODB.Connection
    
    cn.Open "Driver={PostgreSQL ODBC Driver(ANSI)};Server=Localhost;Port=5432;Database=DB_DesafioDoCodigo;Uid=postgres;Pwd=1234;"
    
    If cn.State = adStateOpen Then
        MsgBox "A conexão foi realizada com sucesso!"
    End If
    
    ProcessarImportacaoClientesExcel
    
    
    
    
    
    
    
    
    
    cn.Close
    Set cn = Nothing
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro inesperado no Form_Load, frmPrincipal"
    
End Sub

Private Sub ProcessarImportacaoClientesExcel()
On Error GoTo TrataErro
    
    'Declaração do objeto Common Dialog.
    Dim cd As Object
    Set cd = CreateObject("MSComDlg.CommonDialog")
    
    Dim strCaminhoArquivo As String
    
   
    Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    
    'Processa a importação do arquivo xlsx
    cd.FileName = ""
    cd.Filter = "Arquivos Excel(*.xlsx)|*.xlsx"
    cd.Action = 1

    strCaminhoArquivo = cd.FileName
    blnRegistroEvento = False






    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro inesperado no ProcessarImportacaoClientesExcel, frmPrincipal"
    
End Sub
