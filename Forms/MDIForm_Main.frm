VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.MDIForm MDIForm_Main 
   BackColor       =   &H00FFFFFF&
   Caption         =   "CESNet - Aplicativo de Gerenciamento para o CEJA"
   ClientHeight    =   7455
   ClientLeft      =   225
   ClientTop       =   750
   ClientWidth     =   11265
   Icon            =   "MDIForm_Main.frx":0000
   LinkTopic       =   "MDIForm1"
   Picture         =   "MDIForm_Main.frx":030A
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList IL_LogoEstado 
      Left            =   8460
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":4659B
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSWinsockLib.Winsock Winsock_Main 
      Left            =   9540
      Top             =   120
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList IL_Main16x16 
      Left            =   8700
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8BD6F
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8C08B
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8C3A7
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8C6C3
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8CB17
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8CB75
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8CE91
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8D1AD
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8D4C9
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8D7E3
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm_Main.frx":8F4ED
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar ToolBar_Menu 
      Align           =   1  'Align Top
      Height          =   630
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   1111
      ButtonWidth     =   2143
      ButtonHeight    =   953
      Appearance      =   1
      ImageList       =   "IL_Main16x16"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   9
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dados Pessoais"
            Key             =   "DadosPessoais"
            Object.ToolTipText     =   "Dados Pessoais"
            ImageIndex      =   9
            Object.Width           =   240
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Matricula"
            Key             =   "Matricula"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Trafego"
            Key             =   "EmpMod"
            Object.ToolTipText     =   "Emprestimo de Módulo"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Avaliação"
            Key             =   "Avaliacao"
            Object.ToolTipText     =   "Avaliação"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Orientação"
            Key             =   "Orientacao"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Nota"
            Key             =   "Nota"
            Object.ToolTipText     =   "Lançamento de Notas"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atendimento"
            Key             =   "Atendimento"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Logoff"
            Key             =   "Logof"
            ImageIndex      =   6
         EndProperty
      EndProperty
      Begin VB.PictureBox picMail 
         BorderStyle     =   0  'None
         Height          =   495
         Left            =   10440
         Picture         =   "MDIForm_Main.frx":8F807
         ScaleHeight     =   495
         ScaleWidth      =   495
         TabIndex        =   2
         Top             =   60
         Width           =   495
      End
      Begin MSComDlg.CommonDialog cdMain 
         Left            =   9720
         Top             =   60
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Timer T_StatusBD 
         Interval        =   600
         Left            =   9360
         Top             =   120
      End
   End
   Begin MSComctlLib.StatusBar StatusBar_Menu 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   7080
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   9
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2822
            MinWidth        =   2822
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Text            =   "Usuário:"
            TextSave        =   "Usuário:"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "CAPS"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1270
            MinWidth        =   1270
            TextSave        =   "16:30"
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1931
            MinWidth        =   1940
            TextSave        =   "08/05/2013"
         EndProperty
         BeginProperty Panel7 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   2647
            MinWidth        =   2647
            Text            =   "Versão:"
            TextSave        =   "Versão:"
         EndProperty
         BeginProperty Panel8 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Bevel           =   0
            Object.ToolTipText     =   "Status de Conexão"
         EndProperty
         BeginProperty Panel9 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2824
            MinWidth        =   2824
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MouseIcon       =   "MDIForm_Main.frx":8FB11
   End
   Begin VB.Menu Cadastros 
      Caption         =   "&Cadastros"
      Begin VB.Menu PreMatr 
         Caption         =   "Pre-Matricula"
      End
      Begin VB.Menu DadosPessoais 
         Caption         =   "Dados Pessoais"
      End
      Begin VB.Menu Matricula 
         Caption         =   "Matricula"
      End
      Begin VB.Menu MatriculaAviso 
         Caption         =   "Aviso/Bloqueio de Matrícula"
      End
      Begin VB.Menu Prof 
         Caption         =   "Professor"
      End
      Begin VB.Menu Unidade 
         Caption         =   "Unidade"
      End
      Begin VB.Menu AtendimentosDiversos 
         Caption         =   "Atendimentos Diversos"
      End
      Begin VB.Menu ImportExportMatricula 
         Caption         =   "Importar e Exportar Matricula"
         Begin VB.Menu ImportMatricula 
            Caption         =   "Importar Matricula"
         End
         Begin VB.Menu ExpMatr_RioCard 
            Caption         =   "Exportar Matricula para o RioCard"
         End
      End
      Begin VB.Menu space0 
         Caption         =   "-"
      End
      Begin VB.Menu Ensino 
         Caption         =   "Curso"
      End
      Begin VB.Menu Disciplina 
         Caption         =   "Disciplina"
      End
      Begin VB.Menu Serie 
         Caption         =   "Série"
      End
      Begin VB.Menu Modulo 
         Caption         =   "Módulo"
      End
      Begin VB.Menu Deficiencias 
         Caption         =   "Deficiência"
      End
      Begin VB.Menu InstEnsino 
         Caption         =   "Instituição de Ensino"
      End
      Begin VB.Menu OcorrenciaConclusao 
         Caption         =   "Ocorrência de Conclusão"
      End
   End
   Begin VB.Menu Secretaria 
      Caption         =   "&Secretaria"
   End
   Begin VB.Menu Acervo 
      Caption         =   "&Acervo"
      Begin VB.Menu Trafego 
         Caption         =   "Trafego"
         Begin VB.Menu CadProvas 
            Caption         =   "Cadastro de Provas"
         End
         Begin VB.Menu EmpModulo 
            Caption         =   "Empréstimo de Módulo"
         End
      End
      Begin VB.Menu Biblioteca 
         Caption         =   "Biblioteca"
         Begin VB.Menu BiblCadastroLivro 
            Caption         =   "Cadastro de Livro(s)"
         End
         Begin VB.Menu EmpLivros 
            Caption         =   "Empréstimo de Livro(s)"
         End
      End
   End
   Begin VB.Menu Avaliacao 
      Caption         =   "A&valiação"
      Begin VB.Menu Prova 
         Caption         =   "Prova"
      End
      Begin VB.Menu Nota 
         Caption         =   "Nota"
      End
      Begin VB.Menu space3 
         Caption         =   "-"
      End
      Begin VB.Menu Orientacao 
         Caption         =   "Orientação"
      End
   End
   Begin VB.Menu Relatorios 
      Caption         =   "&Relatórios"
      Begin VB.Menu EstatisticaGeral 
         Caption         =   "Estatística Geral"
         Visible         =   0   'False
      End
      Begin VB.Menu Declaracoes 
         Caption         =   "Declarações"
      End
      Begin VB.Menu Estatistica 
         Caption         =   "Estatística"
         Begin VB.Menu ProvasEfetuadas 
            Caption         =   "Provas Efetuadas"
         End
         Begin VB.Menu Orientacoes 
            Caption         =   "Orientações"
         End
         Begin VB.Menu EstatisticaAlunos 
            Caption         =   "Estatística de Alunos"
         End
      End
      Begin VB.Menu Listagem 
         Caption         =   "Listagem"
         Begin VB.Menu ListagemAlunos 
            Caption         =   "Matrículas por Ensino"
         End
         Begin VB.Menu AlunosCadastrados 
            Caption         =   "Alunos Cadastrados"
         End
         Begin VB.Menu ListagemAlunosDeficientes 
            Caption         =   "Alunos Deficientes"
         End
         Begin VB.Menu ListGradProvas 
            Caption         =   "Grade de Provas"
         End
         Begin VB.Menu ProvasAplicadas 
            Caption         =   "Provas Aplicadas"
         End
         Begin VB.Menu ModulosEmprestados 
            Caption         =   "Modulos Emprestados"
         End
         Begin VB.Menu RAtendimentosDiversos 
            Caption         =   "Atendimentos Diversos"
         End
         Begin VB.Menu RptRetornos 
            Caption         =   "Retornos"
         End
      End
      Begin VB.Menu space 
         Caption         =   "-"
      End
      Begin VB.Menu GerenciadorDeclaracoes 
         Caption         =   "Gerenciador de Declarações"
      End
   End
   Begin VB.Menu Exibir 
      Caption         =   "&Exibir"
      Begin VB.Menu BF 
         Caption         =   "Barra de Ferramentas"
         Checked         =   -1  'True
      End
      Begin VB.Menu BS 
         Caption         =   "Barra de Status"
         Checked         =   -1  'True
      End
      Begin VB.Menu FechaJanelas 
         Caption         =   "Fechar Janelas"
      End
      Begin VB.Menu spc00 
         Caption         =   "-"
      End
      Begin VB.Menu ResultProvas 
         Caption         =   "Resultado das Provas"
      End
   End
   Begin VB.Menu Ferramentas 
      Caption         =   "&Ferramentas"
      Begin VB.Menu Config 
         Caption         =   "Co&nfigurações"
         Begin VB.Menu FundoTela 
            Caption         =   "Fundo de Tela"
            Visible         =   0   'False
         End
         Begin VB.Menu Sistema 
            Caption         =   "Sistema"
         End
         Begin VB.Menu ImprCertificado 
            Caption         =   "Impressão de Certificado"
         End
      End
      Begin VB.Menu Seguranca 
         Caption         =   "Segurança"
         Begin VB.Menu CadUsuarios 
            Caption         =   "Usuários"
         End
         Begin VB.Menu GrupoUsu 
            Caption         =   "Grupo de Usuário"
         End
         Begin VB.Menu UsuTrocarSenha 
            Caption         =   "Trocar Senha"
         End
         Begin VB.Menu UsuRegistroLog 
            Caption         =   "Registro de Log"
         End
      End
      Begin VB.Menu Correio 
         Caption         =   "Correio"
      End
      Begin VB.Menu SQLExecute 
         Caption         =   "SQL Execute"
      End
      Begin VB.Menu spc 
         Caption         =   "-"
      End
      Begin VB.Menu ManutBD 
         Caption         =   "Manutenção do Banco de Dados"
      End
      Begin VB.Menu Ajuste 
         Caption         =   "Ajuste "
         Begin VB.Menu Mensao 
            Caption         =   "Mensão"
         End
         Begin VB.Menu MatrDisciplina 
            Caption         =   "Colocar Dt Inicio em Disciplina"
         End
         Begin VB.Menu TabMatrProva 
            Caption         =   "Tabela de Provas e MatrProva"
         End
         Begin VB.Menu AjustRenovMatr 
            Caption         =   "Ajusta a Dt. de Renovação de Matricula"
         End
      End
   End
   Begin VB.Menu AjudaMenu 
      Caption         =   "Aj&uda"
      Begin VB.Menu Ajuda 
         Caption         =   "Ajuda"
         Shortcut        =   {F1}
      End
      Begin VB.Menu Space_Ajuda0 
         Caption         =   "-"
      End
      Begin VB.Menu SuporteOnLine 
         Caption         =   "Suporte On Line"
      End
      Begin VB.Menu Space_Ajuda1 
         Caption         =   "-"
      End
      Begin VB.Menu SobSistema 
         Caption         =   "Sobre o Sistema"
      End
   End
End
Attribute VB_Name = "MDIForm_Main"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUnidade As Recordset
Dim LogOff As Boolean 'Informa ao sistema se a MDI fecha ou faz logoff
Private Sub Ajuda_Click()
    MsgBox "O banco de dados de ajuda esta ausente."
End Sub

Private Sub AjustRenovMatr_Click()
    If Usuario <> "LEO" Then Exit Sub
    Dim Rst         As Recordset
    Dim sSQL        As String
    Dim DtRenovacao As String
       
    sSQL = "SELECT * FROM MatriculaEnsino " & _
           "WHERE Trancado=FALSE AND DtFinal IS NULL AND DtRenovacao IS NULL AND StatusMatr='ATIVO'"
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Todas as datas preenchidas"
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                DtRenovacao = ChkDtProxRenovacao(Rst.Fields("MatrID"), IIf(IsNull(Rst.Fields("DtInicio")), "", Rst.Fields("DtInicio")))
                Rst.Edit
                Rst.Fields("DtRenovacao") = IIf(DtRenovacao = "00/00/0000", Null, DtRenovacao)
                Rst.Update
                Rst.MoveNext
            Loop
    End If
    Rst.Close
    MsgBox "Fim do processo"
End Sub

Private Sub AlunosCadastrados_Click()
    Form_RelatAlunosCadastrados.Show
End Sub

Private Sub AtendimentosDiversos_Click()
    Form_RegAtendimentos.Show
End Sub

Private Sub BiblCadastroLivro_Click()
    Form_BiblCadastro.Show
End Sub

Private Sub BS_Click()
    If BS.Checked = True Then
            StatusBar_Menu.Visible = False
            BS.Checked = False
        Else
            StatusBar_Menu.Visible = True
            BS.Checked = True
    End If
End Sub
Private Sub BF_Click()
    If BF.Checked = True Then
            ToolBar_Menu.Visible = False
            BF.Checked = False
        Else
            ToolBar_Menu.Visible = True
            BF.Checked = True
    End If
End Sub
Private Sub CadProvas_Click()
    Form_CadProvas.Show
End Sub




Private Sub CadUsuarios_Click()
    Form_Usuario.Show
End Sub



Private Sub Correio_Click()
    Form_Mail.Show
End Sub

Private Sub DadosPessoais_Click()
    Form_DadosPessoais.Show
End Sub

Private Sub Declaracoes_Click()
    Form_Declaracoes.Show
End Sub

Private Sub Deficiencias_Click()
    Form_Deficiencias.Show
End Sub

Private Sub EmpLivros_Click()
    Form_BiblEmpLivro.Show
End Sub

Private Sub EmpModulo_Click()
    Form_EmprModulo.Show
End Sub

Private Sub Ensino_Click()
    Form_Ensino.Show
End Sub

Private Sub EstatisticaAlunos_Click()
    Form_FiltroEstAlunos.Show
End Sub

Private Sub EstatisticaGeral_Click()
Dim Orientacao As String
Orientacao = "SELECT Ensino.Descr, Disciplina.Descr, Count(MatriculaOrientacao.DisciplinaID) AS ContarDeDisciplinaID " & _
"FROM (MatriculaOrientacao INNER JOIN Ensino ON MatriculaOrientacao.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaOrientacao.DisciplinaID = Disciplina.ID " & _
"GROUP BY Ensino.Descr, Disciplina.Descr " & _
"ORDER BY Ensino.Descr, Disciplina.Descr"

Dim mProva As String
mProva = "SELECT Ensino.Descr, Disciplina.Descr, Count(MatriculaProva.MatrID) AS ContarDeMatrID " & _
"FROM (MatriculaProva INNER JOIN Ensino ON MatriculaProva.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaProva.DisciplinaID = Disciplina.ID " & _
"GROUP BY Ensino.Descr, Disciplina.Descr"


Dim x As Recordset

    Set x = BD.OpenRecordset(mProva)
    rptEstatisticaGeral.Sections("Corpo").Controls.Item("Text4").DataSource = x
    rptEstatisticaGeral.Sections("Corpo").Controls.Item("Text4").DataField = x.Fields("ensino.descr")
    
    Call Relatorio(rptEstatisticaGeral, Orientacao)
    rptEstatisticaGeral.Show 1
End Sub

Private Sub ExpMatr_RioCard_Click()
    Form_ExpMatrRioCard.Show
End Sub


Private Sub FechaJanelas_Click()
    Dim Frms As Integer
    Frms = Forms.Count
    Do While Frms > 1
        Unload Forms(Frms - 1)
        If Frms = Forms.Count Then Exit Do
        Frms = Frms - 1
    Loop
End Sub





Private Sub FundoTela_Click()
    Dim RsTMP   As Recordset
    Dim cor     As String
    'Dim Arquivo22 As String
    'Dim Arquivo2 As String
    'On Error GoTo ErroLocate
    
    cdMain.ShowColor
    cor = cdMain.color
    MDIForm_Main.BackColor = cor
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuarioID)
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Erro ao localizar usuário.", vbInformation, "CESNet - Aviso"
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            RsTMP.Edit
            RsTMP.Fields("CorFundo") = cor
            RsTMP.Update
            RsTMP.Close
    End If
        
    'Arquivo2 = FreeFile
    'Open App.path & "\bc.dat" For Output As Arquivo2
    'Print #Arquivo2, cor
    'Close #Arquivo2
    Exit Sub
ErroLocate:
    RegLogErros Err.Number, Err.Description, Me.Caption, 0
    MsgBox Err.Description, vbInformation, Err.Number
    End

End Sub

Private Sub GerenciadorDeclaracoes_Click()
    Form_GerenciadorDeclaracoes.Show
End Sub

Private Sub GrupoUsu_Click()
    Form_UsuGrupo.Show
End Sub

Private Sub ImportMatricula_Click()
    Form_ImportMatricula.Show
End Sub

Private Sub ImprCertificado_Click()
    Form_CoordImpCert.Show
End Sub


Private Sub InstEnsino_Click()
    Form_InstEnsino.Show
End Sub

Private Sub ListagemAlunos_Click()
    Form_RelatAlunos.Show
End Sub

Private Sub ListagemAlunosDeficientes_Click()
    Dim Criterio As String
    Dim alunosAtivos As Boolean
    alunosAtivos = IIf(MsgBox("Listar somente os alunos que estão com algum curso em andamento?", vbYesNo + vbQuestion, App.EXEName) = vbYes, True, False)
    
    Criterio = "SELECT Deficiencias.Descr, Matriculas.MatrID, Matriculas.Nasc, Matriculas.Nome, Ensino.Descr " & _
            "FROM (MatriculaEnsino INNER JOIN (Matriculas INNER JOIN Deficiencias ON Matriculas.DefID = Deficiencias.ID) ON MatriculaEnsino.MatrID = Matriculas.MatrID) INNER JOIN Ensino ON MatriculaEnsino.EnsinoID = Ensino.ID" & _
            IIf(alunosAtivos = True, " WHERE MatriculaEnsino.DtFinal IS NULL ", " ") & _
            "ORDER BY Deficiencias.Descr, Ensino.Descr"
    
    Call Relatorio(rptListaDeficientes, Criterio)
    rptListaDeficientes.Show
End Sub

Private Sub ListGradProvas_Click()
    Form_RelatProvasCad.Show
End Sub

Private Sub ManutBD_Click()
    Form_InfoBD.Show
End Sub

Private Sub Disciplina_Click()
    Form_Disciplina.Show
End Sub


Private Sub MatrDisciplina_Click()
    If Usuario <> "LEO" Then Exit Sub
    Dim RsSerie     As Recordset
    Dim RsDiscTMP   As Recordset
    Dim cont        As Variant
    Dim tot         As Variant

    If MsgBox("Deseja ajustar a tabela de Disciplinas ref. data de inicio?", vbYesNo, "CESNet - Aviso") = vbNo Then Exit Sub


    Set RsDiscTMP = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina ORDER BY DtInicio")
    If RsDiscTMP.BOF And RsDiscTMP.EOF Then
            RsDiscTMP.Close
            Exit Sub
        Else
            RsDiscTMP.MoveLast
            tot = RsDiscTMP.RecordCount
            Me.Caption = "0/" & tot
            RsDiscTMP.MoveFirst
            
            Do Until RsDiscTMP.EOF
                cont = cont + 1
                DoEvents
                Me.Caption = cont & "/" & tot
                Set RsSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & RsDiscTMP.Fields("MatrID") & "'" & _
                            " AND EnsinoID = " & RsDiscTMP.Fields("EnsinoID") & _
                            " AND DisciplinaID = " & RsDiscTMP.Fields("DisciplinaID") & _
                            " ORDER BY SerieID")
                If RsSerie.BOF And RsSerie.EOF Then
                        RsDiscTMP.Edit
                        RsDiscTMP.Fields("DtInicio") = RsDiscTMP.Fields("DtConclusao")
                        RsDiscTMP.Update
                        RsSerie.Close
                    Else
                        
                        RsSerie.MoveFirst
                        RsDiscTMP.Edit
                        RsDiscTMP.Fields("DtInicio") = IIf(IsNull(RsSerie.Fields("DtIni")), RsDiscTMP.Fields("DtConclusao"), RsSerie.Fields("DtIni"))
                        RsDiscTMP.Update
                        RsSerie.Close
                End If
                RsDiscTMP.MoveNext
            Loop
    End If
End Sub

Private Sub Matricula_Click()
    Form_Matricula.Show
End Sub

Private Sub MatriculaAviso_Click()
    Form_MatriculaAviso.Show
End Sub

Private Sub MDIForm_Load()
    
    
    If ValData = False Then
        End
    End If
    'ValidadeSoftware
    'AbrirBD
    LoadStatusBarr
    LogOff = False
    MDIForm_Main.BackColor = PgCorFundo
    ChecarMail
    'Como o aplicativo sera executado DEMO/FULL
    If TipoUsoSoftware = False Then
        MDIForm_Main.Caption = "CESNet - Aplicativo de Gerenciamento para o CEJA   [VERSÃO DE DEMONSTRAÇÃO]"
    End If
    
End Sub
Private Sub ChecarMail()
    If Trim(UsuarioID) = "" Then
        picMail.Visible = False
        Exit Sub
    End If
    Dim RsTMP As Recordset
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Mail WHERE Para = " & UsuarioID & " AND Novo = TRUE")
    If RsTMP.BOF And RsTMP.EOF Then
            picMail.Visible = False
        Else
            picMail.Visible = True
    End If
    RsTMP.Close


End Sub
Private Function ValData() As Boolean
    If Len(Date) <= 8 Then
            MsgBox "Sua DATA está no formato dd/mm/aa." & Chr(13) & "Por favor altere para dd/mm/aaaa no Painel de Controle > Configurações regionais." & Chr(13) & "O sistema será encerrado.", vbInformation, "CESNet - Aviso"
            ValData = False
            Exit Function
        Else
            If Format(Date, "YYYYMMDD") < "20110101" Then
                MsgBox "Data do Equipamento não corresponde a data corrente favor verificar!", vbInformation, "CESNet - Aviso"
                ValData = False
                Exit Function
            End If
            ValData = True
    End If
End Function


Private Sub MDIForm_Resize()
    picMail.left = (MDIForm_Main.Width - picMail.Width) - 250
    End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
    If LogOff = True Then
        Else
            BD.Close
            
            'Unload MDIForm_Main
            End
            
    End If
End Sub

Private Sub Mensao_Click()
    If Usuario <> "LEO" Then Exit Sub
    MsgBox "Leo Funçao Nota no bd foi modificado"
    Exit Sub
    Dim intTot      As Variant
    Dim intCont     As Variant
    Dim strCab      As Variant
    Dim RsMP        As Recordset
    If Usuario <> "LEO" Then Exit Sub
    If MsgBox("Deseja ajustar a tabela de Provas ref. as matriculas para utilizar nova sigla?", vbYesNo, "CESNet - Aviso") = vbNo Then Exit Sub
    Set RsMP = BD.OpenRecordset("SELECT * FROM MatriculaProva")
    If RsMP.BOF And RsMP.EOF Then
            RsMP.Close
            MsgBox "Nenhum Reg"
        Else
            strCab = Me.Caption
            RsMP.MoveLast
            intTot = RsMP.RecordCount
            intCont = 1
            Me.Caption = intCont & "/" & intTot
            RsMP.MoveFirst
            Do Until RsMP.EOF
                DoEvents
                Me.Caption = intCont & "/" & intTot
                RsMP.Edit
                If RsMP.Fields("Aprovado") = True Then
                        RsMP.Fields("Status") = "HB"
                    Else
                        If RsMP.Fields("Nota") = "0" Then
                                RsMP.Fields("Status") = "NH"
                            Else
                                RsMP.Fields("Status") = "NC"
                        End If
                End If
                RsMP.Update
                RsMP.MoveNext
                intCont = intCont + 1
            Loop
            RsMP.Close
            MsgBox "Concluido"
            Me.Caption = strCab
    End If
End Sub

Private Sub Modulo_Click()
    Form_Modulo.Show
End Sub

Private Sub ModulosEmprestados_Click()
    Form_RelatModulosEmprestados.Show
End Sub

Private Sub Nota_Click()
    Form_Notas.Show
End Sub

Private Sub OcorrenciaConclusao_Click()
    Form_OcorrenciaConclusao.Show
End Sub

Private Sub Orientacao_Click()
    Form_Orientacao.Show
End Sub

Private Sub Orientacoes_Click()
    Form_RelatOrientacao.Show
End Sub

Private Sub picMail_Click()
    Form_Mail.Show
End Sub

Private Sub PreMatr_Click()
    Form_PreMatr.Show
End Sub

Private Sub Prof_Click()
    Form_Professores.Show
End Sub
Private Sub Prova_Click()
    Form_Avaliacao.Show
End Sub



Private Sub ProvasAplicadas_Click()
    Form_FiltroProvasAluno.Show

End Sub

Private Sub ProvasEfetuadas_Click()
    Form_FiltroProvasEfetuadas.Show
End Sub


Private Sub RAtendimentosDiversos_Click()
    Form_RelatAtendimentos.Show
End Sub

Private Sub ResultProvas_Click()
    Form_ResultProvas.Show
End Sub

Private Sub RptRetornos_Click()
    Form_RelatRetornos.Show
End Sub

Private Sub Secretaria_Click()
    Form_Secretaria.Show
End Sub

Private Sub Serie_Click()
    Form_Serie.Show
End Sub

Private Sub Sistema_Click()
    Form_Config.Show
End Sub

Private Sub SobSistema_Click()
    Form_About.Show
End Sub



Private Sub SQLExecute_Click()
    Form_SQLExecute.Show
End Sub


Private Sub SuporteOnLine_Click()
    On Error GoTo TRSuporte
    Dim LocSuporte As String
    
    LocSuporte = PathBD & "\Database\Suporte\suporte.exe "
    Shell LocSuporte, vbNormalFocus
    Exit Sub
TRSuporte:
    MsgBox "Erro ao localizar o aplicativo para acesso ao suporte remoto.", vbInformation, "Aviso"
End Sub

Private Sub T_StatusBD_Timer()
    StatusBD
    ChecarMail
End Sub

Private Sub TabMatrProva_Click()
    If Usuario <> "LEO" Then Exit Sub
    Dim RsTrafego   As Recordset
    Dim RsProva     As Recordset
    Dim RsMatrProva As Recordset
    Dim i           As Variant
    Dim cont        As Variant
    
    Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego ORDER BY RefTrafegoID")
    RsTrafego.MoveLast
    cont = RsTrafego.RecordCount
    i = 0
    RsTrafego.MoveFirst
    Do Until RsTrafego.EOF
        i = i + 1
        DoEvents
        Me.Caption = i & " / " & cont
        
        Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID"))
        If RsProva.BOF And RsProva.EOF Then
            Else
                RsProva.MoveFirst
                Do Until RsProva.EOF
                    RsProva.Edit
                    RsProva.Fields("EnsinoID") = RsTrafego.Fields("EnsinoID")
                    RsProva.Fields("DisciplinaID") = RsTrafego.Fields("DisciplinaID")
                    RsProva.Fields("SerieID") = RsTrafego.Fields("SerieID")
                    RsProva.Fields("ModuloID") = RsTrafego.Fields("ModuloID")
                    RsProva.Update
                    RsProva.MoveNext
                Loop
        End If
        Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID"))
        If RsMatrProva.BOF And RsMatrProva.EOF Then
            Else
                RsMatrProva.MoveFirst
                Do Until RsMatrProva.EOF
                    RsMatrProva.Edit
                    RsMatrProva.Fields("EnsinoID") = RsTrafego.Fields("EnsinoID")
                    RsMatrProva.Fields("DisciplinaID") = RsTrafego.Fields("DisciplinaID")
                    'RsMatrProva.Fields("SerieID") = RsTrafego.Fields("SerieID")
                    'RsMatrProva.Fields("ModuloID") = RsTrafego.Fields("ModuloID")
                    RsMatrProva.Update
                    RsMatrProva.MoveNext
                Loop
        End If
        RsTrafego.MoveNext
    Loop
End Sub

Private Sub ToolBar_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
     Select Case Button.Key
        Case "DadosPessoais"
            Form_DadosPessoais.Show
        Case "Matricula"
            Form_Matricula.Show
        Case "EmpMod"
            Form_EmprModulo.Show
        Case "Avaliacao"
            Form_Avaliacao.Show
        Case "Orientacao"
            Form_Orientacao.Show
        Case "Nota"
            Form_Notas.Show
        Case "Atendimento"
            Form_RegAtendimentos.Show
        Case "Logof"
            LogOff = True
            Unload MDIForm_Main
            If Form_Login.LoadLogin = False Then
                    End
                Else
                    MDIForm_Main.Show
            End If
    End Select
End Sub
Private Sub Unidade_Click()
    Form_Unidade.Show
End Sub

Private Sub UsuRegistroLog_Click()
    Form_Log_Leitura.Show
End Sub

Private Sub UsuTrocarSenha_Click()
    Form_TrocaPwd.TrocarSenha (UsuarioID)
End Sub

Private Sub ValidadeSoftware()
    
    If Date >= CDate("30/10/2012") Then
        ToolBar_Menu.Enabled = False
        With MDIForm_Main
            MsgBox "FAVOR SOLICITAR VALIDAÇÃO DO CESNET.", vbCritical, "CESNet - Aviso!"
            End
            MDIForm_Main.Enabled = False
            Unload MDIForm_Main
        End With
        Exit Sub
    End If
End Sub

