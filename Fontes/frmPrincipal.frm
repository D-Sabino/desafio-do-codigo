VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmPrincipal 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Principal"
   ClientHeight    =   6240
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6240
   ScaleWidth      =   8760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdExibeRelatorio 
      Caption         =   "Gerar relatório"
      Height          =   615
      Left            =   240
      TabIndex        =   2
      Top             =   5520
      Width           =   8415
   End
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

'Constantes referentes as colunas do excel.
'----------------------------------------
Private Const mbytColNome = 1
Private Const mbytColEndereco = 2
Private Const mbytColCidade = 3
Private Const mbytColEstado = 4
Private Const mbytColPais = 5
Private Const mbytColTelefone = 6
Private Const mbytColEmail = 7
'----------------------------------------

Private Sub cmdExibeRelatorio_Click()
    geraRelatorio
End Sub

Private Sub cmdImportXLSX_Click()
    importacaoExcel
    exibirDados
End Sub

Private Sub Form_Load()
    On Error GoTo TrataErro
    
    'Inicializando a conexão com o servidor.
    Dim objConsulta As ADODB.Recordset
    gobjBanco.Open "Driver={PostgreSQL ODBC Driver(ANSI)};Server=Localhost;Port=5432;Database=DB_DesafioDoCodigo;Uid=postgres;Pwd=1234;"
    
'    If gobjBanco.State = adStateOpen Then
'        MsgBox "A conexão foi realizada com sucesso!"
'    End If
    
    exibirDados
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro inesperado ao realizar conexão com o servidor. Form_Load, frmPrincipal"
    
End Sub

Private Sub exibirDados()
    On Error GoTo TrataErro

    Dim objConsulta As ADODB.Recordset
    Set objConsulta = New ADODB.Recordset
    
    Dim strSql As String
    strSql = "SELECT * FROM CAD_CLIENTES ORDER BY NOME ASC"
    objConsulta.Open strSql, gobjBanco
    
    'Caso a tabela esteja vazia, encerra o processo.
    If objConsulta.EOF Then Exit Sub
    
    
    ' Limpa o fgClientes antes de preencher com novos dados
    frmPrincipal.fgClientes.Rows = 1
    frmPrincipal.fgClientes.Cols = objConsulta.Fields.Count

    ' Preenche o fgClientes com as colunas da consulta e altera o tamanho das colunas
    For i = 0 To objConsulta.Fields.Count - 1
        frmPrincipal.fgClientes.Col = i
        frmPrincipal.fgClientes.Text = UCase(Left(objConsulta.Fields(i).Name, 1)) & Mid(objConsulta.Fields(i).Name, 2)
        frmPrincipal.fgClientes.ColWidth(i) = objConsulta.Fields(i).DefinedSize * 20
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
    Me.cmdExibeRelatorio.Width = frmPrincipal.fgClientes.Width
    
    objConsulta.Close
    Set objConsulta = Nothing
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro ao exibir dados na grid. exibirDados, frmPrincipal"
    
End Sub


Private Sub importacaoExcel()
    On Error GoTo TrataErro
    
    'Declaração do objeto Common Dialog.
    Dim cd As Object
    Set cd = CreateObject("MSComDlg.CommonDialog")
    
    Dim strCaminhoArquivo As String
    
    'Comunicação
    '-----------------------
    Dim strSql As String
    Dim objConsulta As ADODB.Recordset
    Set objConsulta = New ADODB.Recordset
    '-----------------------
    
    'Campos da tabela:
    '-----------------------
    Dim strNome As String
    Dim strEndereco As String
    Dim strCidade As String
    Dim strEstado As String
    Dim strPais As String
    Dim strTelefone As String
    Dim strEmail As String
    '-----------------------
    
    'Excel:
    '-----------------------
    Dim xl As New Excel.Application
    Dim xlw As Excel.Workbook
    '-----------------------
    
    'Processa a importação do arquivo xlsx
    cd.FileName = ""
    cd.Filter = "Arquivos Excel(*.xlsx)|*.xlsx"
    cd.Action = 1

    strCaminhoArquivo = cd.FileName
    blnRegistroEvento = False

    'Se nenhum arquivo for selecionado, encerra o processo.
    If Trim(strCaminhoArquivo) = "" Then Exit Sub
        
    'Abrir o arquivo do Excel
    Set xlw = xl.Workbooks.Open(strCaminhoArquivo)
        
    lngContReg = 2
    Do While Not Trim(xlw.Application.Cells(lngContReg, 1).Value) = ""
        
        'Limpa os campos
        strNome = ""
        strEndereco = ""
        strCidade = ""
        strEstado = ""
        strPais = ""
        strTelefone = ""
        strEmail = ""
        
        'Preenche as variaveis com as informações do .XLSX
        strNome = xl.Application.Cells(lngContReg, mbytColNome)
        strEndereco = xl.Application.Cells(lngContReg, mbytColEndereco)
        strCidade = xl.Application.Cells(lngContReg, mbytColCidade)
        strEstado = xl.Application.Cells(lngContReg, mbytColEstado)
        strPais = xl.Application.Cells(lngContReg, mbytColPais)
        strTelefone = xl.Application.Cells(lngContReg, mbytColTelefone)
        strEmail = xl.Application.Cells(lngContReg, mbytColEmail)
        
        'Insere as informações do arquivo .XLSX no banco
        strSql = "INSERT INTO CAD_CLIENTES (NOME, ENDERECO, CIDADE, ESTADO, PAIS, TELEFONE, EMAIL)" & _
                 "VALUES ('" & strNome & "', '" & strEndereco & "', '" & strCidade & "', '" & _
                 strEstado & "', '" & strPais & "', '" & strTelefone & "', '" & strEmail & "')"

        objConsulta.Open strSql, gobjBanco

        lngContReg = lngContReg + 1
    Loop
    lngContReg = lngContReg - 1
    
    'Fechando conexões
    '-----------------
    xlw.Close False
    Set xlw = Nothing
    
    xl.Quit
    Set xl = Nothing
    
    Set objConsulta = Nothing
    '-----------------
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro inesperado ao importar arquivo .XLSX. importacaoExcel, frmPrincipal"
    
End Sub

Private Sub geraRelatorio()
On Error GoTo TrataErro
    
    Dim objConsulta As ADODB.Recordset
    Set objConsulta = New ADODB.Recordset
    Dim strSql As String
    
    Dim strNomeRelatorio As String
    strNomeRelatorio = "Relatório de Clientes"
    
    Dim crpe As New CRAXDRT.Application
    Dim report As CRAXDRT.report
    
    
    'Informações do cliente
    '-------------------------------------
    Dim strCodigo As String
    Dim strNome As String
    Dim strEndereco As String
    Dim strCidade As String
    Dim strEstado As String
    Dim strPais As String
    Dim strTelefone As String
    Dim strEmail As String
    '-------------------------------------
    
    strSql = "SELECT * FROM CAD_CLIENTES ORDER BY NOME ASC"
    objConsulta.Open strSql, gobjBanco
    
    If objConsulta.EOF Then
        MsgBox "Não há informações no banco de dados para exibir, realize a importação."
        objConsulta.Close
        
        Exit Sub
    End If
    
    objConsulta.MoveFirst
    
    'Carregando (mouse)
    Screen.MousePointer = vbHourglass
    
    CriarSchema 'Cria schema do arquivo REPORT01.txt
    Open "C:\Windows\Temp\REPORT01.txt" For Output As #1
    Print #1, "CODIGO;NOME;ENDERECO;CIDADE;ESTADO;PAIS;TELEFONE;EMAIL"
    
    Do While Not objConsulta.EOF
        Print #1, objConsulta("ID") & ";" & _
                  objConsulta("NOME") & ";" & _
                  objConsulta("ENDERECO") & ";" & _
                  objConsulta("CIDADE") & ";" & _
                  objConsulta("ESTADO") & ";" & _
                  objConsulta("PAIS") & ";" & _
                  objConsulta("TELEFONE") & ";" & _
                  objConsulta("EMAIL")
                  
        objConsulta.MoveNext
    Loop
    
    Close #1
    
'    Call PadronizarCrystalReport(Me, "rptCliente")

    'Caminho fixo, adaptavel
'    rptCliente.ReportFileName = "C:\Users\felil\Desktop\Projetos\desafio-do-codigo\Report\REPORT01.rpt"
'    rptMovimentacao.WindowState = crptMaximized
'    rptMovimentacao.Action = 1
    
    
'    CrystalReportViewer1.Refresh

    
    Set report = crpe.OpenReport("C:\Users\felil\Desktop\Projetos\desafio-do-codigo\Report\REPORT01.rpt")
'    report.WindowState = crptMaximized
'    report.WindowShowCloseBtn = True
'    report.WindowShowExportBtn = True
'    report.WindowShowGroupTree = False
'    report.WindowShowNavigationCtls = True
'    report.WindowShowPrintBtn = True
'    report.WindowShowPrintSetupBtn = True
'    report.WindowShowRefreshBtn = True
'    report.WindowShowSearchBtn = True
'    report.WindowShowZoomCtl = True
'    report.WindowShowProgressCtls = True
    







    
    Screen.MousePointer = vbNormal


    
    Exit Sub
    Resume

'TRATAMENTO DE ERRO
'----------------------------------------------------------------------------------------
TrataErro:
    MsgBox "Erro na montagem/exibição do relatório. geraRelatorio, frmPrincipal"
    Close #1
'----------------------------------------------------------------------------------------
End Sub

Private Sub CriarSchema()
    Open "C:\Windows\Temp\schema.ini" For Output As #1
    Print #1, "[REPORT01.txt]"
    Print #1, "ColNameHeader = True"
    Print #1, "Format = Delimited(;)"
    Print #1, "MaxScanRows = 25"
    Print #1, "CharacterSet = ANSI"
    Print #1, "Col1= CODIGO char width 100"
    Print #1, "Col2= NOME char width 100"
    Print #1, "Col3= ENDERECO char width 150"
    Print #1, "Col4= CIDADE char width 75"
    Print #1, "Col5= ESTADO char width 50"
    Print #1, "Col6= PAIS char width 50"
    Print #1, "Col7= TELEFONE char width 25"
    Print #1, "Col8= EMAIL char width 100"
    Close #1
    
    Exit Sub
    Resume
    
TrataErro:
    MsgBox "Erro na criação do SCHEMA.INI. CriarSchema, frmPrincipal"
    Close #1
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gobjBanco.Close
    Set gobjBanco = Nothing
End Sub
