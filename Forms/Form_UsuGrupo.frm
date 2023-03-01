VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_UsuGrupo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Gerenciamento de Grupo"
   ClientHeight    =   6255
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8760
   Icon            =   "Form_UsuGrupo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   8760
   Begin VB.CommandButton Bt_AlterarGrupo 
      Caption         =   "&Alterar Grupo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6900
      TabIndex        =   6
      Top             =   900
      Width           =   1695
   End
   Begin VB.CommandButton Bt_NovoGrupo 
      Caption         =   "&Novo Grupo"
      Height          =   375
      Left            =   6900
      TabIndex        =   3
      Top             =   420
      Width           =   1695
   End
   Begin VB.CommandButton Bt_ExcluirGrupo 
      Caption         =   "&Excluir Grupo"
      Enabled         =   0   'False
      Height          =   375
      Left            =   6900
      TabIndex        =   2
      Top             =   1380
      Width           =   1695
   End
   Begin VB.ComboBox Cb_Grupo 
      Height          =   315
      Left            =   600
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   420
      Width           =   6015
   End
   Begin MSComctlLib.ImageList IL_Usu 
      Left            =   7500
      Top             =   2580
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_UsuGrupo.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_UsuGrupo.frx":062E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_UsuGrupo.frx":0A8E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_UsuGrupo.frx":0EEE
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView TrVw_Usu 
      Height          =   5325
      Left            =   135
      TabIndex        =   1
      Top             =   855
      Width           =   6450
      _ExtentX        =   11377
      _ExtentY        =   9393
      _Version        =   393217
      LabelEdit       =   1
      Style           =   5
      ImageList       =   "IL_Usu"
      Appearance      =   1
      Enabled         =   0   'False
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Grupo:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   480
      Width           =   495
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "GERENCIAMENTO DE GRUPO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8775
   End
End
Attribute VB_Name = "Form_UsuGrupo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsGrupoUsuario As Recordset
Dim RsGrupoForm As Recordset
Dim GrupoID As Integer

Private Sub Bt_AlterarGrupo_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub

    
    If Trim(Cb_Grupo.Text) = "" Then Exit Sub
    Select Case left(Bt_AlterarGrupo.Caption, 2)
        Case "&A"
            TrVw_Usu.Enabled = True
            Bt_AlterarGrupo.Caption = "&Gravar Grupo"
            Bt_NovoGrupo.Enabled = False
            Bt_ExcluirGrupo.Enabled = False
            Cb_Grupo.Enabled = False
        Case "&G"
            TrVw_Usu.Enabled = False
            Bt_AlterarGrupo.Caption = "&Alterar Grupo"
            Bt_NovoGrupo.Enabled = True
            Bt_ExcluirGrupo.Enabled = True
            Cb_Grupo.Enabled = True
            ScanTrvw
            CarregarArvore
    End Select
End Sub

Private Sub Bt_ExcluirGrupo_Click()
    If ChkAcesso(Me.Name, "E") = False Then Exit Sub
    If Trim(Cb_Grupo.Text) = "" Then Exit Sub
    
    If GrupoID = 0 Then Exit Sub
    If GrupoID = 1 Then
        MsgBox "Grupo ADMINISTRADOR não pode ser excluido!", vbExclamation, "CESNet - Aviso!"
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir o grupo " & Cb_Grupo.Text & "?", vbInformation + vbYesNo, "CESNet - Exclusão de Grupo") = vbYes Then
        BD.Execute "DELETE * FROM UsuGrupo WHERE GrupoID = " & GrupoID
        BD.Execute "DELETE * FROM UsuGrupoForm WHERE GrupoID = " & GrupoID
        Cb_Grupo.Clear
        MsgBox "Grupo excluido com sucesso", vbInformation, "CESNet - Aviso!"
    End If
End Sub

Private Sub Bt_NovoGrupo_Click()
    Dim NGrupo As String
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub

    NGrupo = UCase(InputBox("Nome do novo grupo?", "Grupo"))
    If Trim(NGrupo) = "" Then
        Exit Sub
    End If
    If Len(NGrupo) <= 3 Then
        MsgBox "Nome do grupo deve ter mais de 3 (tres) caracteres!", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    If IsNumeric(left(NGrupo, 1)) Then
        MsgBox "Nome do grupo não poder ser iniciado por número!", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    If Len(NGrupo) > 30 Then
        MsgBox "Nome do grupo deve ter no maximo 30 (trinta) caracteres!", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    
    If ValidarSoftware("UsuGrupo") = False Then Exit Sub
    
    Set RsGrupoUsuario = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE Nome = '" & NGrupo & "'")
    If RsGrupoUsuario.BOF And RsGrupoUsuario.EOF Then
            RsGrupoUsuario.AddNew
            RsGrupoUsuario.Fields("Nome") = NGrupo
            RsGrupoUsuario.Fields("UsuID") = UsuarioID
            RsGrupoUsuario.Fields("DtHr") = Now
            RsGrupoUsuario.Update
            DoEvents
            Cb_Grupo.AddItem (NGrupo)
            Cb_Grupo.Text = NGrupo
        Else
            MsgBox "Nome de Grupo de Usuarios ja cadastrado.", vbInformation, "CESNet - Aviso!"
    End If
End Sub


Private Sub Cb_Grupo_Click()
    If Trim(Cb_Grupo.Text) = "" Then Exit Sub
    Set RsGrupoUsuario = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE Nome = '" & Trim(Cb_Grupo.Text) & "'")
    If RsGrupoUsuario.BOF And RsGrupoUsuario.EOF Then
            MsgBox "Não foi possivel encontrar Grupo de Usuarios, tente novamente!", vbInformation, "CESNet - Aviso"
            Bt_NovoGrupo.Enabled = True
            Bt_AlterarGrupo.Enabled = False
            Bt_ExcluirGrupo.Enabled = False
            Exit Sub
        Else
            RsGrupoUsuario.MoveFirst
            GrupoID = RsGrupoUsuario.Fields("GrupoID")
            RsGrupoUsuario.Close
            Bt_AlterarGrupo.Enabled = True
            Bt_ExcluirGrupo.Enabled = True
            CarregarArvore
    End If
End Sub

Private Sub Cb_Grupo_DropDown()
    Cb_Grupo.Clear
    Set RsGrupoUsuario = BD.OpenRecordset("SELECT * FROM UsuGrupo ORDER BY Nome ASC")
    If RsGrupoUsuario.BOF And RsGrupoUsuario.EOF Then
        Else
            RsGrupoUsuario.MoveFirst
            Do Until RsGrupoUsuario.EOF
                Cb_Grupo.AddItem RsGrupoUsuario.Fields("Nome")
                RsGrupoUsuario.MoveNext
            Loop
            RsGrupoUsuario.Close
    End If
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
        
    'CarregarGrupos
    CarregarArvore
    
End Sub

Private Sub TrVw_Usu_NodeClick(ByVal Node As MSComctlLib.Node)
    If Len(Node.Key) > 3 Then
        
        Select Case Node.Image
            Case 3
                'If Node.Image = 3 Then
                        Node.Image = 4
                        'MsgBox Node.Image
                    'Else
                    Exit Sub
            Case 4
                        Node.Image = 3
                        'MsgBox Node.Image
                'End If
                Exit Sub
        End Select
    End If
    'Label2.Caption = "KEY: " & Node.Key & " -  IMAGE: " & Node.Image
End Sub
Private Sub Grupo(GrupoID As String, NGrupo As String)
    TrVw_Usu.Nodes.Add , , GrupoID, NGrupo, 1
End Sub
Private Sub subgrupo(GrupoID As String, subGrupoID As String, nSubGrupo As String)
    TrVw_Usu.Nodes.Add GrupoID, tvwChild, subGrupoID, nSubGrupo, 2
    Call Opcoes(subGrupoID)
End Sub
Private Sub Opcoes(subgrupo As String)
    Set RsGrupoForm = BD.OpenRecordset("SELECT * FROM UsuGrupoForm WHERE GrupoID = " & GrupoID & " AND Form = '" & subgrupo & "'") 'Left(subGrupo, 3) & "'")
    If RsGrupoForm.BOF And RsGrupoForm.EOF Then
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "1", "INCLUIR", 4
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "2", "ALTERAR", 4
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "3", "EXCLUIR", 4
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "4", "CONSULTAR", 4
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "5", "IMPRIMIR", 4
        Else
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "1", "INCLUIR", IIf(RsGrupoForm.Fields("Inc") = True, 3, 4)
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "2", "ALTERAR", IIf(RsGrupoForm.Fields("Alt") = True, 3, 4)
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "3", "EXCLUIR", IIf(RsGrupoForm.Fields("Exc") = True, 3, 4)
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "4", "CONSULTAR", IIf(RsGrupoForm.Fields("Cons") = True, 3, 4)
            TrVw_Usu.Nodes.Add subgrupo, tvwChild, subgrupo & "5", "IMPRIMIR", IIf(RsGrupoForm.Fields("Impr") = True, 3, 4)
    End If
    RsGrupoForm.Close
End Sub

Private Sub CarregarArvore()
    DoEvents
    TrVw_Usu.Nodes.Clear
        Call Grupo("CAD", "Cadastro")
            Call subgrupo("CAD", "Form_Matricula", "Matricula")
            Call subgrupo("CAD", "Form_DadosPessoais", "Dados Pessoais")
            Call subgrupo("CAD", "Form_PreMatr", "Pre-Matricula")
            Call subgrupo("CAD", "Form_MatriculaAviso", "Matricula Aviso")
            Call subgrupo("CAD", "Form_ImportMatricula", "Importar Matricula")
            Call subgrupo("CAD", "Form_Professores", "Professor")
            Call subgrupo("CAD", "Form_Unidade", "Unidade")
            Call subgrupo("CAD", "Form_RegAtendimentos", "Atendimentos")
            Call subgrupo("CAD", "Form_Ensino", "Curso")
            Call subgrupo("CAD", "Form_Disciplina", "Disciplina")
            Call subgrupo("CAD", "Form_Serie", "Série")
            Call subgrupo("CAD", "Form_Modulo", "Módulo")
            Call subgrupo("CAD", "Form_Deficiencias", "Deficiencia")
            Call subgrupo("CAD", "Form_InstEnsino", "Instituição de Ensino")
            Call subgrupo("CAD", "Form_OcorrenciaConclusao", "Ocorrencia da Conclusao")
            
            
        
        Call Grupo("SEC", "Secretaria")
            Call subgrupo("SEC", "Form_Secretaria", "Secretaria")
            
        Call Grupo("TRA", "Trafego")
            Call subgrupo("TRA", "Form_CadProvas", "Cadastro de Provas")
            Call subgrupo("TRA", "Form_EmprModulo", "Emprestimo de Módulo")
            'Call subGrupo("ACE", "BIB", "Biblioteca")
        Call Grupo("BIB", "Biblioteca")
            Call subgrupo("BIB", "Form_BiblCadastro", "Cadastro de Livro")
            Call subgrupo("BIB", "Form_EmpLivro", "Emprestimo de Livro")
        
        Call Grupo("AVL", "Avaliação")
            Call subgrupo("AVL", "Form_Avaliacao", "Avaliação")
            Call subgrupo("AVL", "Form_Notas", "Nota")
            Call subgrupo("AVL", "Form_Orientacao", "Orientação")
        
        Call Grupo("RET", "Relatório")
            Call subgrupo("RET", "Form_Declaracoes", "Declarações")
            'Call subGrupo("RET", "Form_FiltroEstAlunos", "Estatistica de Alunos")
            Call subgrupo("RET", "Form_FiltroProvasEfetuadas", "Provas Efetuadas")
            Call subgrupo("RET", "Form_FiltroEstAlunos", "Estatistica de Alunos")
            Call subgrupo("RET", "Form_RelatAlunos", "Listagem de Matriculas")
            Call subgrupo("RET", "Form_RelatAlunosCadastrados", "Listagem de Alunos Cadastrados")
            Call subgrupo("RET", "Form_RelatModulosEmprestados", "Listagem de Emprestimo de Modulos")
            Call subgrupo("RET", "Form_RelatProvasCad", "Listagem da Grade de Provas")
            Call subgrupo("RET", "Form_RelatRetornos", "Listagem da Retornos")
            
            Call subgrupo("RET", "Form_RelatAtendimentos", "Atendimentos Diversos")
            
            Call subgrupo("RET", "Form_GerenciadorDeclaracoes", "Gerenciador de Declarações")
        
        Call Grupo("CFG", "Configurações")
            Call subgrupo("CFG", "Form_Config", "Sistema")
            Call subgrupo("CFG", "Form_Usuario", "Usuário")
            'Call subGrupo("CFG", "DIR", "Direitos")
            Call subgrupo("CFG", "Form_UsuGrupo", "Grupo de Usuarios")
            Call subgrupo("CFG", "Form_TrocaPwd", "Trocar Senha")
            Call subgrupo("CFG", "Form_InfoBD", "Manutenção do Banco de Dados")
            Call subgrupo("CFG", "Form_SQLExecute", "SQL Execute")
            Call subgrupo("CFG", "Form_Log_Leitura", "Leitura de Log")
        Call Grupo("COR", "Correio")
            Call subgrupo("COR", "Form_Mail", "Correio")
End Sub
Private Sub ScanTrvw()
    
    Dim i As Integer
    Dim Imagem As Integer
    Dim Formul As String
    Dim Grupo(99, 5) As String
    Dim cGrupo As Integer
    
    cGrupo = 0
    
    For i = 1 To TrVw_Usu.Nodes.Count

        Imagem = TrVw_Usu.Nodes.Item(i).Image
        If Imagem = 3 Or Imagem = 4 Then
            Formul = TrVw_Usu.Nodes.Item(i).Key
            'If Len(Formul) = 4 Then
            '    MsgBox "CC"
            'End If
            If Len(Formul) > 5 Then
                    If Grupo(cGrupo, 0) <> left(Formul, Len(Formul) - 1) Then
                        cGrupo = cGrupo + 1
                        Grupo(cGrupo, 0) = left(Formul, Len(Formul) - 1)
                    End If
                    
                'Else
                    Grupo(cGrupo, Right(Formul, 1)) = IIf(Imagem = 3, 1, 0)
            End If
        End If
    Next
    Set RsGrupoUsuario = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE GrupoID = " & GrupoID)
    If RsGrupoUsuario.BOF And RsGrupoUsuario.EOF Then
            MsgBox "Erro ao localizar o Grupo de Usuarios. Operação cancelada.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            BD.Execute "DELETE * FROM UsuGrupoForm WHERE GrupoID = " & GrupoID
            RsGrupoUsuario.Close
    End If
    Set RsGrupoForm = BD.OpenRecordset("SELECT * FROM UsuGrupoForm") ' WHERE GrupoID = " & GrupoID)
    For i = 1 To cGrupo
        RsGrupoForm.AddNew
        RsGrupoForm.Fields("GrupoID") = GrupoID
        RsGrupoForm.Fields("Form") = Grupo(i, 0)
        RsGrupoForm.Fields("Inc") = IIf(Grupo(i, 1) = 1, True, False)
        RsGrupoForm.Fields("Alt") = IIf(Grupo(i, 2) = 1, True, False)
        RsGrupoForm.Fields("Exc") = IIf(Grupo(i, 3) = 1, True, False)
        RsGrupoForm.Fields("Cons") = IIf(Grupo(i, 4) = 1, True, False)
        RsGrupoForm.Fields("Impr") = IIf(Grupo(i, 5) = 1, True, False)
        RsGrupoForm.Fields("UsuID") = UsuarioID
        RsGrupoForm.Fields("DtHr") = Now
        RsGrupoForm.Update
    Next
    RsGrupoForm.Close
    
End Sub
