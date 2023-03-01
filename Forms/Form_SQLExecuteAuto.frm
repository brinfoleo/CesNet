VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_SQLExecuteAuto 
   BorderStyle     =   0  'None
   Caption         =   "Manutenção na Base de Dados"
   ClientHeight    =   1485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1215
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   4515
      Begin MSComctlLib.ProgressBar pb 
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   840
         Width           =   4275
         _ExtentX        =   7541
         _ExtentY        =   450
         _Version        =   393216
         Appearance      =   1
      End
      Begin VB.Label lblExec 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   480
         Width           =   4215
      End
      Begin VB.Label lblMsg 
         Height          =   315
         Left            =   180
         TabIndex        =   2
         Top             =   300
         Width           =   4155
      End
   End
End
Attribute VB_Name = "Form_SQLExecuteAuto"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sCom(1000)  As Variant  'string de comando
Dim iCom        As Integer 'contador de comandos
Dim LocalBD     As String

Private Sub Form_Activate()
    DoEvents
    lblMsg.Caption = "Fechando conexão..."
    FecharConexao
    DoEvents
    lblMsg.Caption = "Efetuando cópia de seguranca..."
    Backup
    DoEvents
    lblMsg.Caption = "Abrindo conexão..."
    AbrirConexao
    DoEvents
    lblMsg.Caption = "Executando atualizações na base de dados..."
    AtualizarBaseDados
    DoEvents
    lblMsg.Caption = "Executando atualizações nas Matriculas..."
    ManutencaotbMatriculaEnsino
    DoEvents
    lblMsg.Caption = "Concluido..."
    Unload Me
End Sub

Private Sub AbrirConexao()
    AbrirBD_DAO
End Sub
Private Sub FecharConexao()
    On Error Resume Next
    BD.Close
End Sub

Private Sub AtualizarBaseDados()
    Dim c As Integer
    DoEvents
    CarregarComandos
    pb.Min = 0
    pb.Max = iCom
    For c = 0 To iCom
        DoEvents
        pb.Value = c
        SQLExecute c
    Next

End Sub
Private Sub CarregarComandos()
    iCom = 0
    '#######################################################################################
    '### Versao 5.1.x
    '#######################################################################################
    sCom(iCom) = "CREATE TABLE RegAtendimento (id COUNTER, UsuID NUMERIC, Dt DATE, Hr TEXT(30), Motivo MEMO)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Unidades ADD COLUMN NomeCompleto TEXT (100)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN NumMinDiscipl NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Deficiencias ADD COLUMN Sigla TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE OcorrenciaConclusao ADD COLUMN Sigla TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN Sigla TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaDisciplina ADD COLUMN DtInicio DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Unidades ADD COLUMN CodEscolar TEXT (10)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaOrientacao ADD COLUMN cont COUNTER": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN Numero TEXT (30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN Compl TEXT (30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN NaturalUF TEXT (2)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Unidade ADD COLUMN NomeCompleto TEXT (100)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Disciplina ADD COLUMN Sigla TEXT(4)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE ProvasTMP ADD COLUMN MPID NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN Inativo YESNO": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN NumMinDiscipl NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN StatusMatr TEXT(50)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN NumMinDiscipl NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Deficiencias ADD COLUMN Sigla TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE OcorrenciaConclusao ADD COLUMN Sigla TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN Sigla TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN Raca TEXT(60)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN OpcaoRel TEXT(60)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE InstEnsino ADD COLUMN CidadeRed TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN CartConjugada NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN HB TEXT(60)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN NH TEXT(60)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN FormEstudo TEXT(20)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE BibliotecaIndice ADD COLUMN Localizacao TEXT(20)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE BibliotecaIndice ADD COLUMN N_Inventario TEXT(20)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE BibliotecaIndice ADD COLUMN QtdTotal NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE BibliotecaEmprestimo ADD COLUMN ProfID NUMERIC": iCom = iCom + 1
    sCom(iCom) = "UPDATE Matriculas SET UnidadeID=1 WHERE ISNULL(UnidadeID)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN NumMinDiscipl NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN UsarCidReduzida NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Ensino ADD COLUMN UsarInstSigla NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaAviso ADD COLUMN id AUTOINCREMENT": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaAviso ADD COLUMN Codigo NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaAviso ADD COLUMN DtAvisar DATE": iCom = iCom + 1
    sCom(iCom) = "CREATE TABLE MatriculaRetorno (id COUNTER, UsuarioID NUMERIC, MatrID TEXT(20), DtRetorno DATE, DtHr TEXT(30))": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN TpSang TEXT(10)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN DtNascMae DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN DtNascPai DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaProva ADD COLUMN Obs TEXT(200)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN UsuarioID NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN DtHrSis DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaDisciplina ADD COLUMN UsuarioID NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaDisciplina ADD COLUMN DtHrSis DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaSerie ADD COLUMN UsuarioID NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaSerie ADD COLUMN DtHrSis DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE EmprestimoModulo ADD COLUMN UsuarioIDEmp NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE EmprestimoModulo ADD COLUMN DtHrEmp DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE EmprestimoModulo ADD COLUMN UsuarioIDDev NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE EmprestimoModulo ADD COLUMN DtHrDev DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaRetorno ADD COLUMN TpMov NUMERIC": iCom = iCom + 1
    sCom(iCom) = "UPDATE MatriculaRetorno SET TpMov=1 WHERE TpMov IS NULL": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN DtRenovacao DATE": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN NumConexao TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Matriculas ADD COLUMN NumCenso TEXT(30)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN UsarHistEscImp NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN nmDocHistEsc TEXT(50)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN BloqRenovVencida NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE Config ADD COLUMN FormEstudo TEXT(20)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN UsuarioID NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE MatriculaEnsino ADD COLUMN DtHrSis TEXT(50)": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE EmprestimoModulo ADD COLUMN UsuarioIDDev NUMERIC": iCom = iCom + 1
    sCom(iCom) = "ALTER TABLE EmprestimoModulo ADD COLUMN DtHrSis TEXT(50)": iCom = iCom + 1
    '#######################################################################################

    iCom = iCom - 1

End Sub
Private Sub SQLExecute(i As Integer)
    On Error GoTo trtErroSQL
    Dim sDescricao  As String
    Dim nErro       As String
    lblExec.Caption = sCom(i)
    BD.Execute sCom(i)
    Exit Sub
trtErroSQL:
    sDescricao = Err.Description & " [" & sCom(i) & "]"
    nErro = Err.Number
    
    Select Case nErro
        Case "3010"
            'Tabela ja existe
        Case "3380"
            'Campo ja existe na tabela
        Case "3371"
            'Alteração de tabela ja existente
        Case Else
            'MsgBox "Ops"
            RegLogErros nErro, sDescricao, "Form_SQLExecuteAuto", UsuarioID
    End Select
End Sub
Public Sub CarregarFormulario()
    LoadLocalBD
    LocalBD = PathBD & "\Database\Dados"
    Me.Show 1
    Unload Me
End Sub
Public Sub Backup()
    On Error Resume Next
    Dim nm As String
    nm = Format(Date, "YYYYMMDD")
    FileCopy LocalBD & "\Dados.mdb", LocalBD & "\BD-" & nm & ".bkp"
End Sub
Private Sub ManutencaotbMatriculaEnsino()
    On Error GoTo TrtExc:
    Dim Rst     As ADODB.Recordset
    Dim sSQL    As String
    Dim tReg    As Long
    Dim rCont   As Integer
    
    
    'Pega os dados do Ensino
    Set Rst = New ADODB.Recordset
    sSQL = "SELECT * FROM MatriculaEnsino ORDER BY MatrID"
    
    Rst.Open sSQL, BD, adOpenDynamic
    If Rst.BOF And Rst.EOF Then
            tReg = 0
            Rodape (tReg)
        Else
            tReg = Rst.RecordCount
            Rst.MoveFirst
            Do Until Rst.EOF
                If Rst.Fields("EnsinoID") = 0 Then
                    BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & Rst.Fields("MatrID") & "' AND EnsinoID=0"
                    RgMv "Matr.:" & Rst.Fields("MatrID") & _
                                     " EnsinoID: 0" & _
                                     " Dt.Inicio:" & IIf(IsNull(Rst.Fields("DtInicio")), "  /  /    ", Rst.Fields("DtInicio")) & _
                                     " Dt.Final:" & IIf(IsNull(Rst.Fields("DtFinal")), "  /  /    ", Rst.Fields("DtFinal")) & _
                                     " Local:" & IIf(IsNull(Rst.Fields("Local")), "", Rst.Fields("Local"))
                End If
                rCont = NumCursosCad(Rst.Fields("MatrID"), Rst.Fields("EnsinoID"))
                If rCont >= 2 Then
                    BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & Rst.Fields("MatrID") & "' AND EnsinoID=" & Rst.Fields("EnsinoID") & " AND " & _
                               "DtFinal IS NULL AND Local IS NULL"
                End If
                Rst.MoveNext
                
            Loop
            Rodape (tReg)
    End If
    BD.Close
    MsgBox "Fim do processo!", vbInformation, "Aviso"
    Exit Sub
TrtExc:

    RgMv "Erro em MatriculaEnsino: " & Err.Number & " - " & Err.Description
    Resume Next
End Sub
Public Sub RgMv(Texto As String)
    On Error GoTo TrtErro
    'Registra oq os usuarios do sistema estao fazendo

    'define o ObjPreview filesystem e demais variaveis
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    Dim msg As String
    Dim caminho As String
    caminho = LocalBD & "\" & Replace(App.EXEName, " ", "") & " - " & Format(Date, "yyyymmDD") & ".txt"
    
    'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
'monta informações para gerar a linha da mensagem
    msg = "[" & Now & "] - " & Texto
' inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)
'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
    Resume Next
    
End Sub
Private Sub Rodape(rAnalise As Integer)
    RegLog "#", String(120, "#")
    RegLog "#", "# Finalizado em        : " & Now
    RegLog "#", "# Registros analisados: " & rAnalise
    RegLog "#", String(120, "#")
    RegLog "#", " "
    RegLog "#", " "
End Sub
Private Function NumCursosCad(MatrID As String, Curso As Integer) As Integer
    Dim Rst     As New ADODB.Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & Curso
    Rst.Open sSQL, BD
    If Rst.BOF And Rst.EOF Then
            NumCursosCad = 0
        Else
            NumCursosCad = Rst.RecordCount
            If NumCursosCad >= 2 Then
                Rst.MoveFirst
                Do Until Rst.EOF
                    RegLog "#", NumCursosCad & " - Matr.:" & Rst.Fields("MatrID") & _
                                     " EnsinoID:" & Rst.Fields("EnsinoID") & _
                                     " Dt.Inicio:" & IIf(IsNull(Rst.Fields("DtInicio")), "  /  /    ", Rst.Fields("DtInicio")) & _
                                     " Dt.Final:" & IIf(IsNull(Rst.Fields("DtFinal")), "  /  /    ", Rst.Fields("DtFinal")) & _
                                     " Local:" & IIf(IsNull(Rst.Fields("Local")), "", Rst.Fields("Local"))
                    Rst.MoveNext
                Loop
                 RegLog "#", String(120, "=")
            End If
    End If
    Rst.Close
End Function

