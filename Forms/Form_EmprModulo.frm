VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_EmprModulo 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Emprestimo de Modulo(s)"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10695
   Icon            =   "Form_EmprModulo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   10695
   Begin VB.CommandButton btFoto 
      Caption         =   "&Foto"
      Enabled         =   0   'False
      Height          =   795
      Left            =   5160
      Picture         =   "Form_EmprModulo.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   5100
      Width           =   5415
   End
   Begin VB.Frame Frame2 
      Height          =   5475
      Left            =   60
      TabIndex        =   5
      Top             =   420
      Width           =   4935
      Begin MSFlexGridLib.MSFlexGrid MSFG_Modulos 
         Height          =   5055
         Left            =   120
         TabIndex        =   6
         Top             =   240
         Width           =   4695
         _ExtentX        =   8281
         _ExtentY        =   8916
         _Version        =   393216
         Cols            =   3
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^Módulo                |^Data do Emprestimo |^Data da Devolução"
      End
   End
   Begin VB.Frame Frame1 
      Height          =   4620
      Left            =   5130
      TabIndex        =   0
      Top             =   435
      Width           =   5475
      Begin VB.CheckBox Chk_ModEmprest 
         Caption         =   "Listar Modulo(s) Emprestado(s)"
         Enabled         =   0   'False
         Height          =   435
         Left            =   3840
         TabIndex        =   21
         Top             =   2340
         Width           =   1455
      End
      Begin VB.Frame Frame_Devolucao 
         Height          =   1515
         Left            =   120
         TabIndex        =   13
         Top             =   2940
         Visible         =   0   'False
         Width           =   5235
         Begin VB.CommandButton Bt_Devolver 
            Caption         =   "Devolver"
            Height          =   795
            Left            =   2640
            Picture         =   "Form_EmprModulo.frx":0614
            Style           =   1  'Graphical
            TabIndex        =   20
            Top             =   540
            Width           =   2115
         End
         Begin MSMask.MaskEdBox Meb_DtDevolucao 
            Height          =   315
            Left            =   660
            TabIndex        =   17
            Top             =   780
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "DD/MM/YYYY"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label10 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   16
            Top             =   840
            Width           =   375
         End
         Begin VB.Label Label8 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "DEVOLUÇÃO"
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
            TabIndex        =   14
            Top             =   120
            Width           =   5295
         End
      End
      Begin VB.Frame Frame_Emprestimo 
         Height          =   1515
         Left            =   120
         TabIndex        =   11
         Top             =   2940
         Visible         =   0   'False
         Width           =   5235
         Begin VB.CommandButton Bt_Emprestar 
            Caption         =   "Emprestar"
            Height          =   795
            Left            =   2640
            Picture         =   "Form_EmprModulo.frx":0A56
            Style           =   1  'Graphical
            TabIndex        =   19
            Top             =   540
            Width           =   2115
         End
         Begin MSMask.MaskEdBox Meb_DtEmprestimo 
            Height          =   315
            Left            =   660
            TabIndex        =   18
            Top             =   780
            Width           =   1275
            _ExtentX        =   2249
            _ExtentY        =   556
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "DD/MM/YYYY"
            Mask            =   "##/##/####"
            PromptChar      =   "_"
         End
         Begin VB.Label Label9 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
            Height          =   255
            Left            =   120
            TabIndex        =   15
            Top             =   840
            Width           =   435
         End
         Begin VB.Label Label7 
            Alignment       =   2  'Center
            BackColor       =   &H00FF0000&
            Caption         =   "EMPRESTIMO"
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
            TabIndex        =   12
            Top             =   120
            Width           =   5235
         End
      End
      Begin VB.ComboBox Cb_Disciplina 
         Enabled         =   0   'False
         Height          =   315
         Left            =   960
         TabIndex        =   9
         Top             =   2400
         Width           =   2595
      End
      Begin MSMask.MaskEdBox MebMatricula 
         Height          =   375
         Left            =   240
         TabIndex        =   7
         Top             =   480
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   0
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.####"
         PromptChar      =   "_"
      End
      Begin VB.Label Lb_Ensino 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   960
         TabIndex        =   22
         Top             =   1980
         Width           =   2595
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Curso:"
         Height          =   255
         Left            =   315
         TabIndex        =   10
         Top             =   2040
         Width           =   555
      End
      Begin VB.Label Lb_Nome 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   240
         TabIndex        =   8
         Top             =   1200
         Width           =   4815
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Disciplina:"
         Height          =   255
         Left            =   165
         TabIndex        =   4
         Top             =   2505
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Aluno:"
         Height          =   195
         Left            =   120
         TabIndex        =   3
         Top             =   960
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "Matricula:"
         Height          =   165
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   705
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "EMPRESTIMO DE MÓDULO(S)"
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
      TabIndex        =   2
      Top             =   0
      Width           =   10695
   End
End
Attribute VB_Name = "Form_EmprModulo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatricula As Recordset

Dim RsEnsino As Recordset
Dim RsDisciplina As Recordset
Dim RsModulo As Recordset
Dim RsMaxModulos As Recordset
Dim RsTrafego As Recordset
Dim RsMatriculaDisciplina As Recordset
Dim RsMatriculaSerie As Recordset
Dim RsEmprestimoModulo As Recordset
Dim RsTMP As Recordset

Dim MaxModulos As String

Dim MatrID As String
Dim EnsinoID As Integer
Dim Ensino As String
Dim DisciplinaID As Integer
Dim Disciplina As String
Dim SerieID As Integer
Dim Serie As String
Dim ModuloID As Integer
Dim Modulo As String

Dim lin As Integer
Dim yn As Integer

Private Sub btFoto_Click()
    Form_ExibirImagem.ExibirFoto (MatrID)
End Sub

Private Sub Cb_Disciplina_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Chk_ModEmprest_Click()
    If Chk_ModEmprest.Value = 1 Then
            Cb_Disciplina.Enabled = False
            Cb_Disciplina.Clear
            MSFG_Modulos.Rows = 1
            MSFG_Modulos.Rows = 2
            MSFG_Modulos.Cols = 4
            MSFG_Modulos.FormatString = "^Ensino           |^Disciplina           |^Modulo          |^Emprestimo   "
            Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND ISNULL(DtDevolucao)")
            If RsEmprestimoModulo.BOF And RsEmprestimoModulo.EOF Then
                    MsgBox "Nenhum Modulo foi emprestado!", vbInformation, "CESNet - Aviso!"
                    Chk_ModEmprest.Value = 0
                Else
                    RsEmprestimoModulo.MoveFirst
                    lin = 1
                    Do Until RsEmprestimoModulo.EOF
                        EnsinoID = RsEmprestimoModulo.Fields("EnsinoID")
                        DisciplinaID = RsEmprestimoModulo.Fields("DisciplinaID")
                        ModuloID = RsEmprestimoModulo.Fields("ModuloID")
                        'PegarDescr
                        MSFG_Modulos.TextMatrix(lin, 0) = PgNomeEnsino(EnsinoID)
                        MSFG_Modulos.TextMatrix(lin, 1) = PgNomeDisciplina(DisciplinaID)
                        MSFG_Modulos.TextMatrix(lin, 2) = PgNomeModulo(ModuloID)
                        MSFG_Modulos.TextMatrix(lin, 3) = RsEmprestimoModulo.Fields("DtEmprestimo")
                        RsEmprestimoModulo.MoveNext
                        lin = lin + 1
                        MSFG_Modulos.Rows = MSFG_Modulos.Rows + 1
                    Loop
                    MSFG_Modulos.Rows = MSFG_Modulos.Rows - 1
            End If
        Else
            Cb_Disciplina.Enabled = False
            MSFG_Modulos.Rows = 1
            MSFG_Modulos.Rows = 2
            MSFG_Modulos.Cols = 3
            MSFG_Modulos.FormatString = "^Módulo                |^Data do Emprestimo |^Data da Devolução"
            MstDadosAluno
    End If
End Sub
Private Sub Bt_Devolver_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    If Meb_DtDevolucao.Text = "" Then
        MsgBox "Informe a data.", vbInformation, "CESNet - Aviso!"
        Meb_DtDevolucao.SetFocus
        Exit Sub
    End If
    'PegarID
    Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND ModuloID = " & ModuloID)
    With RsEmprestimoModulo
        .Edit
        .Fields("DtDevolucao") = Meb_DtDevolucao.Text
        .Fields("UsuarioIDDev") = UsuarioID
        .Fields("DtHrDev") = Now()
        .Update
    End With
    Chk_ModEmprest.Value = 0
    Cb_Disciplina_Click
    Frame_Devolucao.Visible = False
End Sub

Private Sub Bt_Emprestar_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    If Meb_DtEmprestimo.Text = "" Then
        MsgBox "Informe a data.", vbInformation, "CESNet - Aviso!"
        Meb_DtEmprestimo.SetFocus
        Exit Sub
    End If
    If ValidarSoftware("EmprestimoModulo") = False Then Exit Sub
    'PegarID
    Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND ModuloID = " & ModuloID)
    With RsEmprestimoModulo
        If .BOF And .EOF Then
            Else
                yn = MsgBox("Deseja Emprestar o Módulo novamente", vbYesNo, "CESNet - Aviso")
                If yn = 6 Then
                        .MoveFirst
                        .Delete
                    Else
                        Cb_Disciplina_Click
                        Frame_Emprestimo.Visible = False
                        Exit Sub
                End If
        End If
        .AddNew
        .Fields("MatrID") = MatrID
        .Fields("EnsinoID") = EnsinoID
        .Fields("DisciplinaID") = DisciplinaID
        .Fields("ModuloID") = ModuloID
        .Fields("DtEmprestimo") = Meb_DtEmprestimo.Text
        .Fields("UsuarioIDEmp") = UsuarioID
        .Fields("DtHrEmp") = Now()
        .Update
    End With
    Cb_Disciplina_Click
    Frame_Emprestimo.Visible = False
End Sub
Private Sub Cb_Disciplina_Click()
    MSFG_Modulos.Rows = 1
    MSFG_Modulos.Rows = 2
    If Cb_Disciplina.Text = "" Then Exit Sub
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina WHERE Descr = '" & Cb_Disciplina.Text & "'")
    If RsDisciplina.BOF And RsDisciplina.EOF Then
        MsgBox "Disciplina Invalida! Por favor, verifique.", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    RsDisciplina.MoveFirst
    DisciplinaID = RsDisciplina.Fields("ID")
    'Pergar Serie do Aluno
    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND Aprovado = false ORDER BY SerieID")
    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
        MsgBox "Nenhuma Serie Localizada. Chame o suporte", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    RsMatriculaSerie.MoveFirst
    
    lin = 1
    yn = 0
    Do Until RsMatriculaSerie.EOF
        SerieID = RsMatriculaSerie.Fields("SerieID")
        Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID & " ORDER BY ModuloID")
        If RsTrafego.BOF And RsTrafego.EOF Then
            MsgBox "Não existe modulo cadstrado para esta Disciplina.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        End If
        RsTrafego.MoveFirst
        Set RsTMP = BD.OpenRecordset("SELECT * FROM Modulo")
        Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
        Do Until RsTrafego.EOF
            ModuloID = RsTrafego.Fields("ModuloID")
            RsTMP.FindFirst "ID = " & ModuloID
            MSFG_Modulos.TextMatrix(lin, 0) = RsTMP.Fields("Descr")
            With RsEmprestimoModulo
                .FindFirst "ModuloID = " & ModuloID
                If .NoMatch Then
                    Else
                        MSFG_Modulos.TextMatrix(lin, 1) = IIf(IsNull(.Fields("DtEmprestimo")), " ", .Fields("DtEmprestimo"))
                        MSFG_Modulos.TextMatrix(lin, 2) = IIf(IsNull(.Fields("DtDevolucao")), " ", .Fields("DtDevolucao"))
                        If MSFG_Modulos.TextMatrix(lin, 1) <> "" And MSFG_Modulos.TextMatrix(lin, 2) = " " Then
                            If yn = 1 Then
                                Else
                                    yn = 1
                            End If
                        End If
                End If
            End With
            MSFG_Modulos.Rows = MSFG_Modulos.Rows + 1
            lin = lin + 1
            RsTrafego.MoveNext
        Loop
        RsMatriculaSerie.MoveNext
    Loop
    MSFG_Modulos.Rows = MSFG_Modulos.Rows - 1
End Sub

Private Sub Cb_Disciplina_DropDown()
 If MebMatricula.Text = "" Then
        Exit Sub
    End If
    Cb_Disciplina.Clear
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND IsNull(DtConclusao)")
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.BOF Then
            MsgBox "Não existe nenhuma Disciplina cadastrada para esta Matricula.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsMatriculaDisciplina.MoveFirst
            Do Until RsMatriculaDisciplina.EOF
                Cb_Disciplina.AddItem (PgNomeDisciplina(RsMatriculaDisciplina.Fields("DisciplinaID")))
                RsMatriculaDisciplina.MoveNext
            Loop
    End If
End Sub


Private Sub Cb_Disciplina_GotFocus()
    Frame_Emprestimo.Visible = False
    Frame_Devolucao.Visible = False
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
    'MaxModulos = 2
    'ValidEmpModulo = True
End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub

Private Sub MebMatricula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        MatrID = formBuscar.IniciarBusca("Matriculas")
        If MatrID = 0 Or Trim(MatrID) = "" Then Exit Sub
        MebMatricula.Text = MatrID
        Chk_ModEmprest.Enabled = True
        Chk_ModEmprest.Value = 0
        MstDadosAluno
    End If

End Sub

Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MatrID = MebMatricula.Text
        Chk_ModEmprest.Enabled = True
        Chk_ModEmprest.Value = 0
        MstDadosAluno
    End If
End Sub
Private Sub MstDadosAluno()
    Frame_Emprestimo.Visible = False
    Frame_Devolucao.Visible = False
    MSFG_Modulos.Rows = 1
    MSFG_Modulos.Rows = 2
    Lb_Nome.Caption = ""

    Lb_Ensino.Caption = ""
    Cb_Disciplina.Clear
    Cb_Disciplina.Enabled = False
    'cb_Ensino.Clear
    btFoto.Enabled = True
   '***** Checar Aviso ******
    If PgAviso(MatrID) = True Then
        Exit Sub
    End If
    '*************************
    MSFG_Modulos.Rows = 1
    MSFG_Modulos.Rows = 2
    EnsinoID = PgMatrEnsino(MatrID, False)
    If EnsinoID = 0 Then
            Lb_Ensino.Caption = ""
            Cb_Disciplina.Clear
            Cb_Disciplina.Enabled = False
            MsgBox "Não existe CURSO cadastrado para esta matricula!", vbInformation, "CESNet - Aviso"
            'Exit Sub
        Else
            Lb_Ensino.Caption = PgNomeEnsino(EnsinoID)
            Cb_Disciplina.Clear
            Cb_Disciplina.Enabled = True
    End If
    
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID = '" & MatrID & "'")
    If RsMatricula.BOF And RsMatricula.EOF Then
        MsgBox "Matricula Invalida!", vbInformation, "CESNet - Aviso!"
        Lb_Nome.Caption = ""
        btFoto.Enabled = False
        MebMatricula.SetFocus
        MebMatricula_GotFocus
        Exit Sub
    End If
    RsMatricula.MoveFirst
    Lb_Nome.Caption = RsMatricula.Fields("Nome")
    'cb_Ensino.Enabled = True
    QtdModulosEmpr
End Sub

Private Sub MSFG_Modulos_Click()
    
    With MSFG_Modulos
        If Chk_ModEmprest.Value = 1 Then
            EnsinoID = PgIDEnsino(.TextMatrix(.Row, 0))
            DisciplinaID = PgIDDisciplina(.TextMatrix(.Row, 1))
            ModuloID = PgIDModulo(.TextMatrix(.Row, 2))
            Frame_Devolucao.Visible = True
            Frame_Emprestimo.Visible = False
            Meb_DtDevolucao.Text = Date
            Exit Sub
        End If
        If .TextMatrix(.Row, 0) = "" Then Exit Sub
        ModuloID = PgIDModulo(.TextMatrix(.Row, 0))
        If .TextMatrix(.Row, 2) = " " Then
                'Devolucao
                Frame_Devolucao.Visible = True
                Frame_Emprestimo.Visible = False
                Meb_DtDevolucao.Text = Date
            Else
                'Emprestimo
                If yn = 1 Then
                    MsgBox "Esta Disciplina ja possue modulo emprestado.", vbInformation, "CESNet - Aviso!"
                    Frame_Devolucao.Visible = False
                    Frame_Emprestimo.Visible = False
                    Exit Sub
                End If
                If QtdModulosEmpr = False Then
                    MsgBox "Esta Matricula ja esta com o numero maximo de modulos emprestados", vbInformation, "CESNet - Aviso!"
                    Exit Sub
                End If
                Frame_Devolucao.Visible = False
                Frame_Emprestimo.Visible = True
                Meb_DtEmprestimo.Text = Date
        End If
    End With
End Sub
Private Function QtdModulosEmpr() As Boolean
    Set RsMaxModulos = BD.OpenRecordset("SELECT * FROM Config")
    If RsMaxModulos.BOF And RsMaxModulos.EOF Then
            Exit Function
        Else
            RsMaxModulos.MoveFirst
            MaxModulos = RsMaxModulos.Fields("MaxModulosEmpr")
    End If
    Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND ISNULL(DtDevolucao)")
    If RsEmprestimoModulo.BOF And RsEmprestimoModulo.EOF Then
            QtdModulosEmpr = True
            Exit Function
        Else
            RsEmprestimoModulo.MoveLast
            If RsEmprestimoModulo.RecordCount = MaxModulos Then
                    QtdModulosEmpr = False
                 Else
                    QtdModulosEmpr = True
            End If
    End If
End Function
