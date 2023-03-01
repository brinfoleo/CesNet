VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_BiblEmpLivro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Emprestimo de Livro(s)"
   ClientHeight    =   7245
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10440
   Icon            =   "Form_BiblEmpLivro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7245
   ScaleWidth      =   10440
   Begin VB.CommandButton btoFiltro 
      Caption         =   "&Buscar Livro"
      Height          =   495
      Left            =   8880
      TabIndex        =   16
      Top             =   1380
      Width           =   1455
   End
   Begin TabDlg.SSTab SST_EmpLivro 
      Height          =   3345
      Left            =   90
      TabIndex        =   6
      Top             =   3825
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   5900
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   520
      TabCaption(0)   =   "Dados do Aluno"
      TabPicture(0)   =   "Form_BiblEmpLivro.frx":030A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Ficha de Emprestimo"
      TabPicture(1)   =   "Form_BiblEmpLivro.frx":0326
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Frame2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Bt_Devolver"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.CommandButton Bt_Devolver 
         Caption         =   "Devolver"
         Enabled         =   0   'False
         Height          =   600
         Left            =   8550
         TabIndex        =   15
         Top             =   1575
         Width           =   1455
      End
      Begin VB.Frame Frame2 
         Height          =   2895
         Left            =   135
         TabIndex        =   13
         Top             =   360
         Width           =   8295
         Begin MSFlexGridLib.MSFlexGrid MSFG_FichaEmp 
            Height          =   2625
            Left            =   135
            TabIndex        =   14
            Top             =   180
            Width           =   7980
            _ExtentX        =   14076
            _ExtentY        =   4630
            _Version        =   393216
            Cols            =   4
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"Form_BiblEmpLivro.frx":0342
         End
      End
      Begin VB.Frame Frame1 
         Height          =   2805
         Left            =   -74865
         TabIndex        =   7
         Top             =   405
         Width           =   9960
         Begin VB.ComboBox Cb_Nome 
            Height          =   315
            Left            =   900
            TabIndex        =   8
            Top             =   585
            Width           =   6225
         End
         Begin MSMask.MaskEdBox Meb_Matricula 
            Height          =   285
            Left            =   900
            TabIndex        =   9
            Top             =   225
            Width           =   1365
            _ExtentX        =   2408
            _ExtentY        =   503
            _Version        =   393216
            Appearance      =   0
            MaxLength       =   11
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.####"
            PromptChar      =   "_"
         End
         Begin VB.Label Lb_Dados 
            Caption         =   "DADOS DO ALUNO"
            Height          =   1005
            Left            =   900
            TabIndex        =   12
            Top             =   945
            Width           =   8340
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome:"
            Height          =   195
            Left            =   315
            TabIndex        =   11
            Top             =   675
            Width           =   510
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Matricula:"
            Height          =   195
            Left            =   135
            TabIndex        =   10
            Top             =   270
            Width           =   690
         End
      End
   End
   Begin VB.ComboBox Cb_Organizar 
      Height          =   315
      ItemData        =   "Form_BiblEmpLivro.frx":03D8
      Left            =   8820
      List            =   "Form_BiblEmpLivro.frx":03E8
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   720
      Width           =   1590
   End
   Begin VB.CommandButton Bt_Emprestar 
      Caption         =   "Emprestar"
      Enabled         =   0   'False
      Height          =   600
      Left            =   8865
      TabIndex        =   3
      Top             =   2025
      Width           =   1455
   End
   Begin VB.Frame Frame3 
      Height          =   3345
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   8700
      Begin MSFlexGridLib.MSFlexGrid MSFG_Livros 
         Height          =   2985
         Left            =   135
         TabIndex        =   2
         Top             =   225
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   5265
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_BiblEmpLivro.frx":0410
      End
   End
   Begin VB.Label Label5 
      Caption         =   "Organizar por:"
      Height          =   195
      Left            =   8820
      TabIndex        =   5
      Top             =   495
      Width           =   1005
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "EMPRESTIMO DE LIVRO(S)"
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
      Width           =   10440
   End
End
Attribute VB_Name = "Form_BiblEmpLivro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatricula As Recordset
Dim RsBiblIndice As Recordset
Dim RsBiblAssunto As Recordset
Dim RsBiblEmprestimos As Recordset
Dim RsDisciplina As Recordset

Dim DisciplinaID As String

Dim MatrID As String
Dim LivroID As String

Dim tmp As String

Private Sub Bt_Devolver_Click()
    If MSFG_FichaEmp.TextMatrix(MSFG_FichaEmp.Row, 3) <> " " Then
        MsgBox "Livro ja devolvido.", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    Set RsBiblEmprestimos = BD.OpenRecordset("SELECT * FROM BibliotecaEmprestimo WHERE LivroID = " & LivroID) ' & "'")
    If RsBiblEmprestimos.BOF And RsBiblEmprestimos.EOF Then
            MsgBox "Erro ao localizar Aluno ", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsBiblEmprestimos.MoveFirst
            'MatrID = RsBiblEmprestimos.Fields("MatrID")
            'MstDadosAluno (MatrID)
            If MsgBox("Confirma DEVOLUÇÃO?" & Chr(13) & _
                        "Matricula: " & MatrID & Chr(13) & _
                        "Livro: " & LivroID & Chr(13) & _
                        "Titulo: " & PgTituloLivro(LivroID) & Chr(13) & _
                        "Data Devolucao: " & Date _
                        , vbYesNo, "CESNet - Devolucao de Livro(s)") = vbYes Then
                Set RsBiblEmprestimos = BD.OpenRecordset("SELECT * FROM BibliotecaEmprestimo WHERE LivroID = " & LivroID & "AND IsNull(DtDevolucao)")
                If RsBiblEmprestimos.BOF And RsBiblEmprestimos.EOF Then
                        MsgBox "Erro ao localizar Aluno ", vbInformation, "CESNet - Aviso"
                        Exit Sub
                    Else
                        With RsBiblEmprestimos
                            '.MoveFirst
                            .Edit
                            .Fields("DtDevolucao") = "21/11/1978"
                            .Update
                        End With
                End If
                Set RsBiblIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE LivroID = " & LivroID)
                With RsBiblIndice
                    .MoveFirst
                    .Edit
                    .Fields("Emprestado") = False
                    '.Fields("DtEmprestimo") = ""
                    .Update
                End With
            End If
    End If
    Call MstFichaEmp(MatrID)
    Call Cb_Organizar_Click
End Sub






Private Sub btoFiltro_Click()
    Dim strFiltro As String
    
   LivroID = Form_BiblFiltro.CarregarFormulario
   If LivroID = 0 Then Exit Sub
   strFiltro = " WHERE LivroID = " & LivroID
   FiltrarLivros strFiltro
   'LKX 8534 Fiat Verde
   
End Sub

Private Sub Meb_Matricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID ='" & Meb_Matricula.Text & "'")
        With RsMatricula
            If .BOF And .EOF Then
                    MsgBox "Matricula não encontrada.", vbInformation, "CESNet - Aviso!"
                    'LimpDados
                    Meb_Matricula.SetFocus
                    Exit Sub
                Else
                    MatrID = Meb_Matricula.Text
                    MstDadosAluno (MatrID)
                    'LstHstEscolar
            End If
        End With
    End If
End Sub
Private Sub Meb_Matricula_GotFocus()
    Meb_Matricula.SelStart = 0
    Meb_Matricula.SelLength = 11
End Sub



Private Sub Bt_Emprestar_Click()
    If MatrID = "" Then
        MsgBox "Erro ao localizar a Matricula do Aluno Por favor Verifique!", vbInformation, "CESNet - AVISO"
        Exit Sub
    End If
    If MSFG_Livros.TextMatrix(MSFG_Livros.Row, 5) = "SIM" Then
        MsgBox "Livro ja emprestado", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
   
    If MsgBox("Confirma EMPRESTIMO?" & Chr(13) & _
        "Matricula: " & MatrID & Chr(13) & _
        "Livro: " & LivroID & Chr(13) & _
        "Titulo: " & PgTituloLivro(LivroID) & Chr(13) & _
        "Data Emprestimo: " & Date _
        , vbYesNo, "CESNet - Emprestimo de Livro(s)") = vbYes Then
            Set RsBiblIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE LivroID = " & LivroID)
            With RsBiblIndice
                .MoveFirst
                .Edit
                .Fields("Emprestado") = True
                .Fields("DtEmprestimo") = Date
                .Update
            End With
            Set RsBiblEmprestimos = BD.OpenRecordset("SELECT * FROM BibliotecaEmprestimo")
            With RsBiblEmprestimos
                .AddNew
                .Fields("MatrID") = MatrID
                .Fields("LivroID") = LivroID
                .Fields("DtEmprestimo") = Date
                .Update
            End With
            MsgBox "Livro " & LivroID & " - " & PgTituloLivro(LivroID) & ", emprestado com sucesso!", vbInformation, "CESNet - Emprestimo de Livro(s)"
            
        Else
            'MsgBox "nao"
    End If
    Call MstFichaEmp(MatrID)
    Call Cb_Organizar_Click
End Sub

Private Sub Cb_Nome_Click()
    If Cb_Nome.Text = "" Then Exit Sub
    With RsMatricula
        .FindFirst "Nome ='" & Cb_Nome.Text & "'"
        If .NoMatch Then
                MsgBox "Erro no acesso ao Banco de Dados." & Chr(13) & "Por favor, reinicie o formulário!", vbExclamation, "aviso!"
                Unload Me
            Else
                MstDadosAluno (.Fields("MatrID"))
                'LstHstEscolar
        End If
    End With
End Sub
Private Sub Cb_Nome_DropDown()
    'If acao = 1 Or acao = 2 Then Exit Sub
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE Nome LIKE '" & Cb_Nome.Text & "*' ORDER BY Nome")
    If RsMatricula.BOF And RsMatricula.EOF Then
            'Cb_Nome.Clear
        Else
            LstNomeAlunos
    End If
End Sub

Private Sub Cb_Nome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    If Len(Cb_Nome.Text) = 50 Then
        KeyAscii = 0
        Beep
    End If
End Sub

Private Sub LstNomeAlunos()
    With RsMatricula
        If .BOF And .EOF Then
                Exit Sub
            Else
                .MoveFirst
                tmp = Cb_Nome.Text
                Cb_Nome.Clear
                Cb_Nome.Text = tmp
                Do While .EOF = False
                    Cb_Nome.AddItem (.Fields("Nome"))
                    .MoveNext
                Loop
        End If
    End With
End Sub
Private Function MstDadosAluno(Matricula As String)
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID = '" & Matricula & "'")
    If RsMatricula.BOF And RsMatricula.EOF Then
            MsgBox "Erro ao localizar Aluno", vbInformation, "CESNet - AVISO"
            Exit Function
        Else
            RsMatricula.MoveFirst
    End If
    With RsMatricula
        MatrID = .Fields("MatrID")
        Meb_Matricula.Text = MatrID
        Cb_Nome.Text = .Fields("Nome")
        Lb_Dados.Caption = .Fields("Unidade") & Chr(13) & _
                            IIf(Trim(.Fields("End")) = "", "<Não Cadastrado>", .Fields("End")) & Chr(13) & _
                            .Fields("Bai") & " - " & .Fields("Mun") & "/" & .Fields("UF") & Chr(13) & _
                            .Fields("Tel1") & "/" & .Fields("Tel2") & Chr(13) & _
                            IIf(Trim(.Fields("RG")) = "", "<ALUNO NÃO HABILITADO, POIS NÃO CADASTROU O RG.>", "RG: " & .Fields("RG"))
    End With
    MstFichaEmp (MatrID)
End Function
Private Sub MstFichaEmp(MatrID As String)
    Set RsBiblEmprestimos = BD.OpenRecordset("SELECT * FROM BibliotecaEmprestimo WHERE MatrID = '" & MatrID & "' ORDER BY Seq")
    If RsBiblEmprestimos.BOF And RsBiblEmprestimos.EOF Then
            MSFG_FichaEmp.Rows = 1
            MSFG_FichaEmp.Rows = 2
        Else
            With MSFG_FichaEmp
                .Rows = 1
                RsBiblEmprestimos.MoveFirst
                Do Until RsBiblEmprestimos.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = Mid(String(5, "0"), 1, 5 - Len(RsBiblEmprestimos.Fields("LivroID"))) & RsBiblEmprestimos.Fields("LivroID")
                    .TextMatrix(.Rows - 1, 1) = PgTituloLivro(RsBiblEmprestimos.Fields("LivroID"))
                    .TextMatrix(.Rows - 1, 2) = RsBiblEmprestimos.Fields("DtEmprestimo")
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(RsBiblEmprestimos.Fields("DtDevolucao")), " ", RsBiblEmprestimos.Fields("DtDevolucao"))
                    RsBiblEmprestimos.MoveNext
                Loop
            End With
    End If
End Sub
Private Function PgTituloLivro(LivroID As String)
    LivroID = Mid(String(5, "0"), 1, 5 - Len(LivroID)) & LivroID
    Set RsBiblIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE LivroID = " & LivroID)
    If RsBiblIndice.BOF And RsBiblIndice.EOF Then
            Exit Function
        Else
            RsBiblIndice.MoveFirst
            PgTituloLivro = RsBiblIndice.Fields("Titulo")
    End If
End Function
Private Sub Cb_Organizar_Click()
    Select Case Cb_Organizar.Text
        Case "Disciplina"
            FiltrarLivros (" ORDER BY " & Cb_Organizar.Text & "ID")
        Case Else
            FiltrarLivros (" ORDER BY " & Cb_Organizar.Text)
    End Select
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    SST_EmpLivro.Tab = 0
    Cb_Organizar.Text = "LivroID"
End Sub
Private Function FiltrarLivros(Filtro As String)
    Set RsBiblIndice = BD.OpenRecordset("SELECT * FROM BibliotecaIndice " & Filtro)
    If RsBiblIndice.BOF And RsBiblIndice.EOF Then
            MSFG_Livros.Rows = 1
            MSFG_Livros.Rows = 2
        Else
            With RsBiblIndice
                .MoveFirst
                MSFG_Livros.Rows = 1
                Do Until .EOF
                    DoEvents
                    MSFG_Livros.Rows = MSFG_Livros.Rows + 1
                    LivroID = Mid(String(4, "0"), 1, 4 - Len(.Fields("LivroID"))) & .Fields("LivroID") '& .Fields("Prateleira")
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 0) = LivroID
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 1) = .Fields("Titulo")
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 2) = .Fields("Autor")
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 3) = PgNomeDisciplina(.Fields("DisciplinaID"))
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 4) = IIf(IsNull(.Fields("Assunto")), "", .Fields("Assunto"))
                    If .Fields("Emprestado") = True Then
                        MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 5) = "SIM"
                        MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 6) = IIf(IsNull(.Fields("DtEmprestimo")), " ", .Fields("DtEmprestimo"))
                        MSFG_Livros.Col = 1
                        MSFG_Livros.Row = MSFG_Livros.Rows - 1
                        MSFG_Livros.ColSel = MSFG_Livros.Cols - 1
                        MSFG_Livros.FillStyle = flexFillRepeat
                        MSFG_Livros.CellForeColor = &HFF&
                    End If
                    .MoveNext
                Loop
            End With
    End If
End Function

Private Sub MSFG_FichaEmp_Click()
    LivroID = MSFG_FichaEmp.TextMatrix(MSFG_FichaEmp.Row, 0)
    If MSFG_FichaEmp.TextMatrix(MSFG_FichaEmp.Row, 3) <> " " Then
            Bt_Devolver.Enabled = False
        Else
            Bt_Devolver.Enabled = True
    End If
    
    
End Sub

Private Sub MSFG_Livros_Click()
    LivroID = MSFG_Livros.TextMatrix(MSFG_Livros.Row, 0)
    If MSFG_Livros.TextMatrix(MSFG_Livros.Row, 5) = "SIM" Then
            Bt_Emprestar.Enabled = False
            'Bt_Devolver.Enabled = True
        Else
            Bt_Emprestar.Enabled = True
            'Bt_Devolver.Enabled = False
    End If
End Sub
