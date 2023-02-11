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
    
    'Inicializando a conexão com o servidor.
    Dim cn As New ADODB.Connection
    Dim ConectaBanco As String
    

'    ConectaBanco = "DNS=Postgre30;Database=DB_DesafioDoCodigo;Server=Localhost;Uid=postgres;Port=5432;pwd=1234;"
'    cn.Open ConectaBanco
'
    
    
    
    
    
    cn.Open "Driver={PostgreSQL ODBC Driver(ANSI)};Server=Localhost;Port=5432;Database=DB_DesafioDoCodigo;Uid=postgres;Pwd=1234;"
    
    If cn.State = adStateOpen Then
        MsgBox "A conexão foi realizada com sucesso!"
    End If
    
    cn.Close
    Set cn = Nothing
    
End Sub
