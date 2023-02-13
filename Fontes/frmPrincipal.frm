VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Principal"
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6540
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin MSFlexGridLib.MSFlexGrid fgClientes 
      Height          =   4455
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   7858
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
   End
   Begin VB.CommandButton cmdImportXLSX 
      Caption         =   "Importe o arquivo .XLSX"
      Height          =   615
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "frmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public gobjBanco As New ADODB.Connection

Private Sub cmdImportXLSX_Click()
    importacaoExcel
    frmPrincipal.Refresh
End Sub

Private Sub Form_Load()
    On Error GoTo TrataErro
    
    'Inicializando a conexão com o servidor.
    Dim objConsulta As ADODB.Recordset
    Dim strSQL As String
    
    gobjBanco.Open "Driver={PostgreSQL ODBC Driver(ANSI)};Server=Localhost;Port=5432;Database=DB_DesafioDoCodigo;Uid=postgres;Pwd=1234;"
    
'    If gobjBanco.State = adStateOpen Then
'        MsgBox "A conexão foi realizada com sucesso!"
'    End If
    
    exibirDados
        
    
    
    gobjBanco.Close
    Set gobjBanco = Nothing
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro inesperado no Form_Load, frmPrincipal"
    
End Sub

Private Sub exibirDados()
    Dim objConsulta As ADODB.Recordset
    Set objConsulta = New ADODB.Recordset
    
    Dim strSQL As String
    strSQL = "SELECT * FROM CAD_CLIENTES ORDER BY NOME ASC"
    
    objConsulta.Open strSQL, gobjBanco
    
    ' Limpa o fgClientes antes de preencher com novos dados
    frmPrincipal.fgClientes.Rows = 1
    frmPrincipal.fgClientes.Cols = objConsulta.Fields.Count

    ' Preenche o fgClientes com as colunas da consulta e altera o tamanho das colunas
    For i = 0 To objConsulta.Fields.Count - 1
        frmPrincipal.fgClientes.Col = i
        frmPrincipal.fgClientes.Text = UCase(Left(objConsulta.Fields(i).Name, 1)) & Mid(objConsulta.Fields(i).Name, 2)
        frmPrincipal.fgClientes.ColWidth(i) = objConsulta.Fields(i).DefinedSize * 10
    Next
    
    ' Preenche o fgClientes com as linhas da consulta
    frmPrincipal.fgClientes.Rows = 2
    frmPrincipal.fgClientes.FixedRows = 1
    objConsulta.MoveFirst
    
    Do Until objConsulta.EOF
        For i = 0 To objConsulta.Fields.Count - 1
            frmPrincipal.fgClientes.Col = i
            frmPrincipal.fgClientes.Row = frmPrincipal.fgClientes.Rows - 1
            frmPrincipal.fgClientes.Text = objConsulta.Fields(i).Value
        Next
        frmPrincipal.fgClientes.Rows = frmPrincipal.fgClientes.Rows + 1
        objConsulta.MoveNext
    Loop
    
    ' Ajusta o tamanho do formulário de acordo com a largura da grid
    Me.Width = frmPrincipal.fgClientes.Width + (Me.Width - Me.ScaleWidth) + 550

    ' Ajusta a largura do botão de acordo com a largura da grid
    Me.cmdImportXLSX.Width = frmPrincipal.fgClientes.Width
    
    objConsulta.Close
    Set objConsulta = Nothing
    
End Sub


Private Sub importacaoExcel()
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

    If Trim(strCaminhoArquivo) = "" Then Exit Sub
        
    'Abrir o arquivo do Excel
    Set xlw = xl.Workbooks.Open(strCaminhoArquivo)
        
    lngContReg = 1
    Do While Not Trim(xlw.Application.Cells(lngContReg, mbytColArea).Value) = ""
        lngContReg = lngContReg + 1
    Loop
    lngContReg = lngContReg - 1
        
    xlw.Close False
    Set xlw = Nothing
    Set xl = Nothing
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro inesperado no importacaoExcel, frmPrincipal"
    
End Sub

