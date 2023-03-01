Attribute VB_Name = "modImportacao"
Option Explicit
Public conexao             As ADODB.Connection

Public Function ImportOpenBD() As Boolean
    On Error GoTo ImportErro
    Dim sConexao As String
    'caminho = "127.0.0.1:C:\Documents and Settings\Leonardo\Desktop\CESNet\BD_CES\CES SENAI\SACES\SACE - CES SENAI\BANCO\SACES.FDB"
    Set conexao = New ADODB.Connection
    
    '#################################################################################################
    '### String de conexao para Firebird do CEJA Senai
    '###
'    sConexao = "DRIVER=Firebird/InterBase(r) driver; UID=SYSDBA; PWD=masterkey;DBNAME=" & ForaLocBD ' & caminho
    sConexao = pgStringConexao
    
    conexao.Open sConexao
    ImportOpenBD = True
    Exit Function
ImportErro:
    MsgBox Err.Description, vbCritical, Err.Number
    ImportOpenBD = False
End Function
Private Function pgStringConexao() As String
    On Error GoTo trtErroString
    Dim Arquivo     As String
    Dim linha       As String
    Arquivo = FreeFile
    Open App.path & "\CESNet.ext" For Input As Arquivo
    Line Input #Arquivo, linha
    Close #Arquivo
    pgStringConexao = Trim(linha)
    Exit Function
trtErroString:
    MsgBox Err.Description, vbCritical, Err.Number
    pgStringConexao = ""
End Function
