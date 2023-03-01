VERSION 5.00
Begin VB.Form Form_BiblCadastro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Biblioteca"
   ClientHeight    =   4770
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8100
   Icon            =   "Form_BiblCadastro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   4770
   ScaleWidth      =   8100
   Begin VB.Frame Frame4 
      Height          =   3255
      Left            =   6030
      TabIndex        =   12
      Top             =   360
      Width           =   1995
      Begin VB.CommandButton Bt_PesqLivro 
         Caption         =   "Pesquisar Livro"
         Height          =   465
         Left            =   90
         TabIndex        =   19
         Top             =   1575
         Width           =   1815
      End
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cancelar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   90
         TabIndex        =   18
         Top             =   2655
         Width           =   1815
      End
      Begin VB.CommandButton Bt_Gravar 
         Caption         =   "Gravar"
         Enabled         =   0   'False
         Height          =   465
         Left            =   90
         TabIndex        =   17
         Top             =   2205
         Width           =   1815
      End
      Begin VB.CommandButton Bt_AlterarLivro 
         Caption         =   "Alterar Livro"
         Height          =   465
         Left            =   90
         TabIndex        =   16
         Top             =   675
         Width           =   1815
      End
      Begin VB.CommandButton Bt_ExcluirLivro 
         Caption         =   "Excluir Livro"
         Height          =   465
         Left            =   90
         TabIndex        =   14
         Top             =   1125
         Width           =   1815
      End
      Begin VB.CommandButton Bt_IncluirLivro 
         Caption         =   "Incluir Livro"
         Height          =   465
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   1815
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Assunto(s):"
      Height          =   1545
      Left            =   90
      TabIndex        =   10
      Top             =   3150
      Width           =   5865
      Begin VB.TextBox Txt_Assunto 
         Enabled         =   0   'False
         Height          =   1185
         Left            =   135
         MaxLength       =   100
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   21
         Top             =   270
         Width           =   5595
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1455
      Left            =   90
      TabIndex        =   4
      Top             =   1620
      Width           =   5865
      Begin VB.TextBox Txt_Autor 
         Enabled         =   0   'False
         Height          =   285
         Left            =   810
         TabIndex        =   15
         Top             =   585
         Width           =   4875
      End
      Begin VB.TextBox Txt_Titulo 
         Enabled         =   0   'False
         Height          =   285
         Left            =   810
         TabIndex        =   11
         Top             =   180
         Width           =   4920
      End
      Begin VB.ComboBox Cb_Disciplinas 
         Enabled         =   0   'False
         Height          =   315
         Left            =   810
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   990
         Width           =   3030
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Titulo:"
         Height          =   195
         Left            =   270
         TabIndex        =   8
         Top             =   225
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Autor:"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   630
         Width           =   510
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Disciplina:"
         Height          =   195
         Left            =   45
         TabIndex        =   6
         Top             =   1035
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1185
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   5865
      Begin VB.ComboBox Cb_Livro 
         Height          =   315
         Left            =   900
         TabIndex        =   20
         Top             =   270
         Width           =   1365
      End
      Begin VB.ComboBox Cb_Prateleira 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form_BiblCadastro.frx":030A
         Left            =   900
         List            =   "Form_BiblCadastro.frx":035C
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   675
         Width           =   1365
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Prateleira:"
         Height          =   195
         Left            =   90
         TabIndex        =   3
         Top             =   720
         Width           =   690
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Livro Nº:"
         Height          =   195
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   645
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "CADASTRO DE LIVROS"
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
      Width           =   8115
   End
End
Attribute VB_Name = "Form_BiblCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBibliotecaIndice As Recordset
Dim RsBibliotecaAssunto As Recordset
Dim RsDisciplina As Recordset

Dim Opcao As Integer '0 - Nenhum // 1 - Incluir // 2 - Alterar

Dim LivroID As String
Dim LivroIDOld As String
Dim DisciplinaID As String

Dim LivroPesq As String
'Dim L1 As String
'Dim L2 As String
Private Sub Bt_AlterarLivro_Click()
    If Txt_Titulo.Text = "" Then
        MsgBox "Por favor selecione um livro", vbInformation, "Aviso"
        Exit Sub
    End If
    Opcao = 2
    LivroIDOld = Cb_Livro.Text '& Cb_Prateleira.Text
    HDForm (True)
    HDBt (False)
    Cb_Livro.Enabled = False
End Sub

Private Sub Bt_Cancelar_Click()
    HDBt (True)
    HDForm (False)
    LimpForm
    Opcao = 0
End Sub


Private Sub Bt_ExcluirLivro_Click()
    If LivroID = 0 Then Exit Sub
    If MsgBox("Deseja realmente excluir este livro?", vbInformation + vbYesNo, "Aviso") = vbYes Then
        BD.Execute "DELETE * FROM BibliotecaIndice WHERE LivroId = " & LivroID
        RegLog "00.000.0000", "Exclusao do livro: " & LivroID & "-" & Txt_Titulo.Text
        LimpForm
        'Cb_Prateleira.Clear
        MsgBox "Livro excluido!", vbInformation, "Aviso"
    End If
End Sub

Private Sub Bt_Gravar_Click()
    If ValidarSoftware("BibliotecaIndice") = False Then Exit Sub
    If ValidaDados = True Then
        GravarLivro
        HDBt (True)
        HDForm (False)
        If Opcao = 1 Then
                MsgBox "Referencia: " & LivroID & Chr(13) & "Livro cadastrado com sucesso!", vbInformation, "Biblioteca"
            Else
                MsgBox "Referencia: " & LivroID & Chr(13) & "Livro alterado com sucesso!", vbInformation, "Biblioteca"
        End If
        Opcao = 0
    End If
End Sub

Private Sub Bt_IncluirLivro_Click()
    Opcao = 1
    LimpForm
    HDForm (True)
    HDBt (False)
End Sub

Private Sub Bt_PesqLivro_Click()
    
    LivroPesq = InputBox("Digire a referencia do livro?", "CESNet - Pesquisa Biblioteca")
    If Trim(LivroPesq) = "" Then Exit Sub
    If Len(LivroPesq) > 5 Then
        MsgBox "Numero de referencia muito grande." & Chr(13) & "Pesquisa cancelada.", vbInformation, "CESNet - Pesquisa"
        Exit Sub
    End If
    LivroPesq = Mid(String(4, "0"), 1, 4 - Len(LivroPesq)) & LivroPesq
    'L1 = Left(LivroPesq, 4)
    'L2 = UCase(Right(LivroPesq, 1))
    If IsNumeric(LivroPesq) Then
                If IsNumeric(left(LivroPesq, 1)) Xor IsNumeric(Right(LivroPesq, 1)) Then
                    
                    MsgBox "Referencia invalida.", vbInformation, "CESNet - Aviso"
                    Exit Sub
                    Else
                End If
            Else
                MsgBox "Referencia invalida.", vbInformation, "CESNet - Aviso"
                Exit Sub
    End If
    'If IsNumeric(L2) Then
    '    MsgBox "Referencia invalida.", vbInformation, "CESNet - Aviso"
    '    Exit Sub
    'End If
    
    Set RsBibliotecaIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE LivroID = " & LivroPesq)
    With RsBibliotecaIndice
        If .BOF And .EOF Then
            MsgBox "Livro não encontrado", vbInformation, "Aviso"
            Exit Sub
        End If
        .MoveLast
        If .RecordCount >= 2 Then
            MsgBox "Erro na pesquisa pois possuem dois livros com a mesma referencia por favor avise ao suporte", vbCritical, "CESNet - Aviso"
            Exit Sub
        End If
        .MoveFirst
        LivroID = LivroPesq
        ExibirLivro (LivroID)
    End With
End Sub


Private Sub Cb_Disciplinas_DropDown()
    Cb_Disciplinas.Clear
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina ORDER BY Descr")
    If RsDisciplina.BOF And RsDisciplina.EOF Then
            MsgBox "Nenhuma Disciplina cadastrada", vbInformation, "CESNet - AVISO"
            Exit Sub
        Else
            With RsDisciplina
                .MoveFirst
                Do Until .EOF
                    Cb_Disciplinas.AddItem (.Fields("Descr"))
                    .MoveNext
                Loop
            End With
    End If

End Sub

Private Sub Cb_Disciplinas_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub



Private Sub Cb_Livro_Click()
    LivroID = Cb_Livro.Text
    ExibirLivro (LivroID)
End Sub

Private Sub Cb_Livro_DropDown()
    Set RsBibliotecaIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice ORDER BY LivroID")
    If RsBibliotecaIndice.BOF And RsBibliotecaIndice.EOF Then
            Cb_Livro.Clear
        Else
            Cb_Livro.Clear
            RsBibliotecaIndice.MoveFirst
            Do Until RsBibliotecaIndice.EOF
                Cb_Livro.AddItem (Mid(String(4, "0"), 1, 4 - Len(RsBibliotecaIndice.Fields("LivroID"))) & RsBibliotecaIndice.Fields("LivroID"))
                RsBibliotecaIndice.MoveNext
            Loop
    End If
End Sub

Private Sub Form_Load()
    Opcao = 0
End Sub
Private Function ExibirLivro(LvID As String)
    If LvID = "" Then
        Exit Function
    End If
    Set RsBibliotecaIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE LivroID = " & LvID)
    If RsBibliotecaIndice.BOF And RsBibliotecaIndice.EOF Then
        Exit Function
    End If
    With RsBibliotecaIndice
        Txt_Titulo.Text = .Fields("Titulo")
        Txt_Autor.Text = .Fields("Autor")
        Cb_Livro.Text = LivroID
        Cb_Prateleira.Text = .Fields("Prateleira")
        'DisciplinaID = .Fields("DisciplinaID")
        Cb_Disciplinas.AddItem (PgNomeDisciplina(.Fields("DisciplinaID")))
        Cb_Disciplinas.Text = PgNomeDisciplina(.Fields("DisciplinaID"))
        Txt_Assunto.Text = IIf(IsNull(.Fields("Assunto")), "", .Fields("Assunto"))
    End With
End Function









Private Sub HDForm(Opcao As Double)
    Cb_Prateleira.Enabled = Opcao
    Cb_Livro.Enabled = IIf(Opcao = True, False, True)
    Txt_Titulo.Enabled = Opcao
    Txt_Autor.Enabled = Opcao
    Cb_Disciplinas.Enabled = Opcao
    Txt_Assunto.Enabled = Opcao
End Sub
Private Sub LimpForm()
    Cb_Prateleira.Text = "A"
    Cb_Livro.Clear
    Txt_Titulo.Text = ""
    Txt_Autor.Text = ""
    Cb_Disciplinas.Clear
    Txt_Assunto.Text = ""
End Sub
Private Sub HDBt(Opcao As Double)
    Bt_IncluirLivro.Enabled = Opcao
    Bt_AlterarLivro.Enabled = Opcao
    Bt_ExcluirLivro.Enabled = Opcao
    Bt_PesqLivro.Enabled = Opcao
    
    Bt_Gravar.Enabled = IIf(Opcao = True, False, True)
    Bt_Cancelar.Enabled = IIf(Opcao = True, False, True)
End Sub
Private Sub GravarLivro()
    Select Case Opcao
        Case 1
            Set RsBibliotecaIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice")
            RsBibliotecaIndice.AddNew
        Case 2
            Set RsBibliotecaIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE LivroID = " & LivroIDOld)
            If RsBibliotecaIndice.BOF And RsBibliotecaIndice.EOF Then
                    MsgBox "Erro ao localizar referencia antiga por favor tente novamente", vbInformation, "CESNet - Aviso"
                Else
                    RsBibliotecaIndice.Edit
            End If
        Case Else
            Exit Sub
    End Select
    With RsBibliotecaIndice
        .Fields("Prateleira") = Cb_Prateleira.Text
        .Fields("Titulo") = IIf(Trim(Txt_Titulo.Text) = "", Null, Trim(Txt_Titulo.Text))
        .Fields("Autor") = IIf(Trim(Txt_Autor.Text) = "", Null, Trim(Txt_Autor.Text))
        .Fields("DisciplinaID") = PgIDDisciplina(Cb_Disciplinas.Text)
        .Fields("Assunto") = IIf(Trim(Txt_Assunto.Text) = "", Null, Trim(Txt_Assunto.Text))
        .Update
        Select Case Opcao
            Case 1
                .MoveLast
                LivroID = Mid(String(4, "0"), 1, 4 - Len(.Fields("LivroID"))) & .Fields("LivroID")
                Cb_Livro.Text = LivroID
        
                'LivroID = LivroID & .Fields("Prateleira")
            Case 2
                LivroID = Mid(String(4, "0"), 1, 4 - Len(Cb_Livro.Text)) & Cb_Livro.Text
                'LivroID = LivroID & Cb_Prateleira.Text
            Case Else
                Exit Sub
        End Select
        
    End With
End Sub
Private Function ValidaDados() As Double
    If Trim(Txt_Titulo.Text) = "" Then
        MsgBox "O campo TITULO não pode ser deixado em branco. Por favor verifique!", vbInformation, "CESNet - Aviso"
        ValidaDados = False
        Exit Function
    End If
    If Trim(Txt_Autor.Text) = "" Then
        MsgBox "O campo AUTOR não pode ser deixado em branco. Por favor verifique!", vbInformation, "CESNet - Aviso"
        ValidaDados = False
        Exit Function
    End If
    If Trim(Cb_Disciplinas.Text) = "" Then
        MsgBox "O campo DISCIPLINA não pode ser deixado em branco. Por favor verifique!", vbInformation, "CESNet - Aviso"
        ValidaDados = False
        Exit Function
    End If
    If Trim(Cb_Prateleira.Text) = "" Then
        MsgBox "O campo PRATELEIRA não pode ser deixado em branco. Por favor verifique!", vbInformation, "CESNet - Aviso"
        ValidaDados = False
        Exit Function
    End If
    ValidaDados = True
End Function
Private Sub Txt_Assunto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_Autor_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Sub Txt_Titulo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
