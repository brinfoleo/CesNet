VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Begin VB.Form Form_Avaliacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Avaliação"
   ClientHeight    =   6000
   ClientLeft      =   135
   ClientTop       =   360
   ClientWidth     =   10155
   Icon            =   "Form_Avaliacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6000
   ScaleWidth      =   10155
   Begin VB.CommandButton btFoto 
      Caption         =   "&Foto"
      Enabled         =   0   'False
      Height          =   915
      Left            =   8520
      Picture         =   "Form_Avaliacao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5040
      Width           =   1515
   End
   Begin VB.TextBox Txt_Obs 
      Appearance      =   0  'Flat
      ForeColor       =   &H000000FF&
      Height          =   555
      Left            =   630
      MaxLength       =   200
      MultiLine       =   -1  'True
      TabIndex        =   20
      Top             =   4410
      Width           =   9420
   End
   Begin VB.CommandButton Bt_ImpFolhaResp 
      Caption         =   "&Folha Resposta"
      Enabled         =   0   'False
      Height          =   915
      Left            =   5280
      Picture         =   "Form_Avaliacao.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   5040
      Width           =   1515
   End
   Begin VB.CommandButton Bt_AplicarProva 
      Caption         =   "&Aplicar Prova"
      Enabled         =   0   'False
      Height          =   915
      Left            =   6900
      Picture         =   "Form_Avaliacao.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   5040
      Width           =   1515
   End
   Begin VB.TextBox Txt_Tipo 
      Alignment       =   2  'Center
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1380
      MaxLength       =   1
      TabIndex        =   11
      Top             =   5160
      Width           =   615
   End
   Begin VB.Frame Frame2 
      Height          =   2415
      Left            =   45
      TabIndex        =   3
      Top             =   1485
      Width           =   10035
      Begin MSFlexGridLib.MSFlexGrid MSFG_Provas 
         Height          =   2160
         Left            =   90
         TabIndex        =   5
         Top             =   180
         Width           =   9885
         _ExtentX        =   17436
         _ExtentY        =   3810
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_Avaliacao.frx":0C28
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1095
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   10035
      Begin VB.ComboBox Cb_Disciplina 
         Enabled         =   0   'False
         Height          =   315
         Left            =   6975
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   225
         Width           =   2700
      End
      Begin MSMask.MaskEdBox MebMatricula 
         Height          =   375
         Left            =   855
         TabIndex        =   6
         Top             =   225
         Width           =   1635
         _ExtentX        =   2884
         _ExtentY        =   661
         _Version        =   393216
         Appearance      =   0
         ForeColor       =   -2147483630
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
         ForeColor       =   &H00000000&
         Height          =   315
         Left            =   3150
         TabIndex        =   17
         Top             =   225
         Width           =   2985
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Curso:"
         Height          =   195
         Left            =   2550
         TabIndex        =   16
         Top             =   315
         Width           =   555
      End
      Begin VB.Label Lb_Nome 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   855
         TabIndex        =   9
         Top             =   660
         Width           =   8835
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   225
         TabIndex        =   8
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Disciplina:"
         Height          =   210
         Left            =   6240
         TabIndex        =   2
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Matricula:"
         Height          =   195
         Left            =   45
         TabIndex        =   1
         Top             =   315
         Width           =   735
      End
   End
   Begin MSMask.MaskEdBox Meb_Dt 
      Height          =   615
      Left            =   2160
      TabIndex        =   15
      Top             =   5160
      Width           =   2235
      _ExtentX        =   3942
      _ExtentY        =   1085
      _Version        =   393216
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   18
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
   Begin VB.Label Label8 
      Alignment       =   1  'Right Justify
      Caption         =   "Obs:"
      Height          =   195
      Left            =   225
      TabIndex        =   19
      Top             =   4455
      Width           =   375
   End
   Begin VB.Label Lb_Prova 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
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
      ForeColor       =   &H000000FF&
      Height          =   315
      Left            =   630
      TabIndex        =   14
      Top             =   4050
      Width           =   9435
   End
   Begin VB.Label Label6 
      Caption         =   "Prova:"
      Height          =   195
      Left            =   135
      TabIndex        =   13
      Top             =   4125
      Width           =   495
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Informe o Tipo:"
      Height          =   195
      Left            =   180
      TabIndex        =   10
      Top             =   5400
      Width           =   1095
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "AVALIAÇÃO"
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
      Left            =   -30
      TabIndex        =   4
      Top             =   0
      Width           =   10230
   End
End
Attribute VB_Name = "Form_Avaliacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatricula             As Recordset
Dim RsMatriculaEnsino       As Recordset
Dim RsMatriculaDisciplina   As Recordset
Dim RsMatriculaProvas       As Recordset
Dim RsTrafego               As Recordset
'Dim RsDisciplina As Recordset
Dim RsEmprestimoModulo      As Recordset
Dim RsProvas                As Recordset
Dim RsProvasTMP             As Recordset
Dim RsTMP                   As Recordset
Dim RsOrientacao            As Recordset

Dim EnsinoID                As Integer
Dim DisciplinaID            As Integer
Dim ModuloID                As Integer
Dim SerieID                 As Integer
Dim RefTrafegoID            As Integer

Dim MatrID                  As String
Dim nProva                  As String
Dim nmModulo                As String

Dim lin                     As Integer

Dim Orientacao              As Boolean

Private Sub Bt_ImpFolhaResp_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    
    If MebMatricula.Text = "" Then
        Bt_ImpFolhaResp.Enabled = False
    End If
    If Trim(Lb_Prova.Caption) = "" Then
        Bt_ImpFolhaResp.Enabled = False
    End If
    Call iFolhaResp
End Sub

Private Sub btFoto_Click()
    Form_ExibirImagem.ExibirFoto (MatrID)
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
    Orientacao = False
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Config")
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Dados de Configuração vazio, favor checar.", vbInformation, "CESNet - Atenção"
            Exit Sub
        Else
            RsTMP.MoveFirst
            VincModulo = RsTMP.Fields("VincModulos")
            RsTMP.Close
    End If
End Sub

Private Function ChkData(Dt As String) As Boolean
    Dim tmp As Integer
    Dim d, m, a As String
    d = left(Dt, 2)
    m = Mid(Dt, 4, 2)
    a = Right(Dt, 4)
    If d = 0 Or d > 31 Then
        ChkData = False
        Exit Function
    End If
    If m = 0 Or m > 12 Then
        ChkData = False
        Exit Function
    End If
 
    If a < Right(Date, 4) Or a > Right(Date, 4) Then
        If MsgBox("O ano digitado (" & a & ")  não é o ano corrente (" & Right(Date, 4) & "). Deseja manter?", vbInformation + vbYesNo, "CESNet - Aviso") = vbNo Then
            ChkData = False
            Exit Function
        End If
    End If
    ChkData = True
End Function

Private Sub Bt_AplicarProva_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
  
    If Lb_Nome.Caption = "" Or Txt_Tipo.Text = "" Or Meb_Dt.Text = "" Then
        Exit Sub
    End If
    If ChkData(Meb_Dt.Text) = False Then
        MsgBox "Favor checar o campo data!", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    RefTrafegoID = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 6)
    'Nao deixar o aluno fazer as provas repitidas
    'Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "' ORDER BY Seq, NProva")
    'RsProvasTMP.FindFirst "Tipo = '" & Txt_Tipo.Text & "'"
    'If RsProvasTMP.NoMatch Then
    '    Else
    '        MsgBox "A prova de tipo " & Txt_Tipo.Text & " ja foi aplicada." & Chr(13) & "Por favor aplique outra prova", vbInformation ,"CESNet - Aviso!"
    '        Exit Sub
    'End If
    
    'RsTrafego.FindFirst "ReftrafegoID"
    If ValidarSoftware("MatriculaProva") = False Then Exit Sub
    
    Set RsMatriculaProvas = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0) & "'")
    With RsMatriculaProvas
        If .BOF And .EOF Then
                .AddNew
            Else
                .Edit
        End If
        .Fields("MatrID") = MatrID
        '.Fields("RefTrafegoID") = RefTrafegoID
        .Fields("RefTrafegoID") = RefTrafegoID
        .Fields("EnsinoID") = EnsinoID
        .Fields("DisciplinaID") = DisciplinaID
        .Fields("NProva") = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0)
        .Fields("Assunto") = IIf(Trim(MSFG_Provas.TextMatrix(MSFG_Provas.Row, 2)) = "", Null, MSFG_Provas.TextMatrix(MSFG_Provas.Row, 2))
        .Fields("Tipo") = Txt_Tipo.Text
        .Fields("Nota") = Null
        .Fields("Status") = "NC"
        .Fields("DtAvaliacao") = Meb_Dt.Text
        .Fields("UsuarioIDAv") = UsuarioID
        .Fields("DtHrAv") = Now
        .Update
        RegLog MatrID, "PROVA APLICADA: " & PgNomeEnsino(EnsinoID) & "/" & PgNomeDisciplina(DisciplinaID) & "/" & _
               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0) & "/" & Txt_Tipo.Text & "/" & _
               Meb_Dt.Text & "/" & UsuarioID
    End With
    
            '=============================================================================================
            'FUNCAO TEMPORARIA PARA PEGAR O INICIO DA DISCIPLINA
            Dim RsMatrSerieTMP      As Recordset
            Dim RsMatrDisciplTMP    As Recordset
            Dim DtInicio            As String
            
            Set RsMatrSerieTMP = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY SerieID")
            If RsMatrSerieTMP.BOF And RsMatrSerieTMP.EOF Then
                    RsMatrSerieTMP.Close
                    DtInicio = Meb_Dt.Text
                Else
                    RsMatrSerieTMP.MoveFirst
                    ''Debug.Print RsMatrSerieTMP.Fields("SerieID")
                    DtInicio = IIf(IsNull(RsMatrSerieTMP.Fields("DtIni")), Meb_Dt.Text, RsMatrSerieTMP.Fields("DtIni"))
                    RsMatrSerieTMP.Close
            End If
            
            Set RsMatrDisciplTMP = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
            If RsMatrDisciplTMP.BOF And RsMatrDisciplTMP.EOF Then
                    RsMatrDisciplTMP.Close
                Else
                    RsMatrDisciplTMP.MoveFirst
                    RsMatrDisciplTMP.Edit
                    RsMatrDisciplTMP.Fields("DtInicio") = DtInicio
                    RsMatrDisciplTMP.Update
                    RsMatrDisciplTMP.Close
            End If
            '================================================================================================

    
    Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND isnull(DtInicio)")
    If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
        Else
        With RsMatriculaEnsino
            .Edit
            .Fields("DtInicio") = Meb_Dt.Text
            .Update
        End With
    End If
    MSFG_Provas.TextMatrix(MSFG_Provas.Row, 3) = Meb_Dt.Text
    MSFG_Provas.TextMatrix(MSFG_Provas.Row, 4) = Txt_Tipo.Text
    If MsgBox("Deseja imprimir FOLHA RESPOSTA?", vbYesNo + vbInformation, "CESNet - Atenção") = vbYes Then
        If ChkAcesso(Me.Name, "I") = True Then
            iFolhaResp
        End If
    End If
    'If VincModulo = True Then
    '        MstProvas
    '    Else
    '        MstTodasProvas
    'End If
    
     
    Bt_AplicarProva.Enabled = False
    Lb_Prova.Caption = ""
    Txt_Obs.Text = ""
    Txt_Tipo.Text = ""
    Meb_Dt.PromptInclude = False
    Meb_Dt.Text = ""
    Meb_Dt.PromptInclude = True
    Txt_Tipo.Enabled = False
    Meb_Dt.Enabled = False
    
    
End Sub

Private Sub Cb_Disciplina_Click()
    Lb_Prova.Caption = ""
    Txt_Obs.Text = ""
    Txt_Tipo.Text = ""

    '********Checar Aviso
    If PgAviso(MatrID) = True Then
        MSFG_Provas.Rows = 1
        Exit Sub
    End If
    '*******************
    DisciplinaID = PgIDDisciplina(Cb_Disciplina.Text) 'RsDisciplina.Fields("DisciplinaID")

'Checa se o aluno esta na lista de Orientacao para esta disciplina
    Set RsOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY Cont") ' & " AND HabDisciplina = False")
    If RsOrientacao.BOF And RsOrientacao.EOF Then
            Txt_Tipo.Enabled = True
            Meb_Dt.Enabled = True
            Bt_AplicarProva.Enabled = True
            Orientacao = False
        Else
            
            'RsOrientacao.MoveFirst
            'If RsOrientacao.Fields("HabDisciplina") = False Then
            'nao deixar o aluno proceguir
            RsOrientacao.MoveLast
            If RsOrientacao.Fields("HabDisciplina") = False Then
                    MsgBox "Matricula ultrapassou o número minimo de reprovações.", vbInformation, "CESNet - Aviso"
                    MSFG_Provas.Rows = 1
                    MSFG_Provas.Rows = 2
                    Lb_Prova.Caption = ""
                    Txt_Obs.Text = ""
                    Txt_Tipo.Text = ""
                    Txt_Tipo.Enabled = False
                    Meb_Dt.Enabled = False
                    Bt_AplicarProva.Enabled = False
                    Orientacao = True
                Else
                    Txt_Tipo.Enabled = True
                    Meb_Dt.Enabled = True
                    Bt_AplicarProva.Enabled = True
                    Orientacao = False
            End If
            'End If
    End If
    
    If VincModulo = True Then
            MstProvas
        Else
            MstTodasProvas
    End If
End Sub
Private Sub Meb_Dt_GotFocus()
    Meb_Dt.SelStart = 0
    Meb_Dt.SelLength = 10
End Sub

Private Sub MebMatricula_Change()
    If IsNumeric(MebMatricula.Text) Then
        MatrID = MebMatricula.Text
        MstDadosAluno
        MSFG_Provas.Rows = 1
        MSFG_Provas.Rows = 2
        Bt_AplicarProva.Enabled = False
    End If
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
       MstDadosAluno
    End If
End Sub

Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MatrID = MebMatricula.Text
        MstDadosAluno
        MSFG_Provas.Rows = 1
        MSFG_Provas.Rows = 2
    End If
End Sub
Private Sub MstDadosAluno()
    

    
    '***** Checar Aviso ******
    If PgAviso(MatrID) = True Then
        Lb_Ensino.Caption = ""
        Cb_Disciplina.Clear
        Cb_Disciplina.Enabled = False
        MSFG_Provas.Rows = 1
        MSFG_Provas.Rows = 2
        MSFG_Provas.Enabled = False
        
        Exit Sub
    End If
    '*************************
   
    Orientacao = False 'Status de Orientação
    '27/02/13 - Alterado pois se o curso estivesse trancado ele pegava as
    'disciplinas do mesmo
    EnsinoID = PgMatrEnsino(MatrID, False)
    
    Lb_Ensino.Caption = ""
    Lb_Prova.Caption = ""
    Txt_Obs.Text = ""
    Txt_Tipo.Text = ""
     MSFG_Provas.Enabled = True
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
    
    Cb_Disciplina.Enabled = True
    btFoto.Enabled = True
    If VincModulo = True Then
            Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND EnsinoID=" & EnsinoID & " AND ISNULL(DtDevolucao)")
            With RsEmprestimoModulo
                If .BOF And .EOF Then
                    MsgBox "Não exite registro de provas para esta matricula. Por favor, verifique na Trafego.", vbInformation, "CESNet - Aviso!"
                    Cb_Disciplina.Clear
                    MSFG_Provas.Rows = 1
                    MSFG_Provas.Rows = 2
                    MebMatricula.SetFocus
                    Exit Sub
                End If
                .MoveFirst
                Cb_Disciplina.Clear
                'EnsinoID = .Fields("EnsinoID")
                Lb_Ensino.Caption = PgNomeEnsino(EnsinoID)
                Do Until .EOF
                    Cb_Disciplina.AddItem (PgNomeDisciplina(.Fields("DisciplinaID")))
                    'Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino WHERE EnsinoID = " & EnsinoID)
                    'If RsEnsino.BOF And RsEnsino.EOF Then
                    '        MsgBox "Descrição de ENSINO não encontrado.", vbInformation ,"CESNet - Aviso!"
                    '        Exit Sub
                    '    Else
                    '        Lb_Ensino.Caption = RsEnsino.Fields("Descr")
                    'End If
                    'Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina WHERE DisciplinaID = " & DisciplinaID)
                    'RsDisciplina.MoveFirst
                    'Cb_Disciplina.AddItem (RsDisciplina.Fields("Descr"))
                    .MoveNext
                Loop
            End With
        Else
            Cb_Disciplina.Clear
            EnsinoID = PgMatrEnsino(MatrID)
            Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND IsNull(DtConclusao)")
            If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                    MsgBox "Nenhuma Disciplina cadastrada para esta matricula.", vbInformation, "Atenção"
                    Exit Sub
                Else
                    With RsMatriculaDisciplina
                        .MoveFirst
                        
                        Lb_Ensino.Caption = PgNomeEnsino(EnsinoID)
                        Do Until .EOF
                            Cb_Disciplina.AddItem (PgNomeDisciplina(.Fields("DisciplinaID")))
                            .MoveNext
                        Loop
                    End With
            End If
    End If
        
    
End Sub

Private Sub MSFG_Provas_Click()
    With MSFG_Provas
        If .Row = 1 Then
                'Txt_Tipo.SetFocus
            Else
                Lb_Prova.ForeColor = &HFF&
                Lb_Prova.Caption = .TextMatrix(.Row, 0) & " - " & .TextMatrix(.Row, 3)
                Meb_Dt.PromptInclude = False
                Meb_Dt.Text = .TextMatrix(.Row, 3)
                Meb_Dt.PromptInclude = True
                Txt_Tipo.Text = .TextMatrix(.Row, 4)
                Bt_AplicarProva.Enabled = False
                Txt_Tipo.Enabled = False
                Meb_Dt.Enabled = False
                If Trim(.TextMatrix(.Row, 4)) <> "" And Trim(.TextMatrix(.Row, 5)) = "" Then
                        Bt_ImpFolhaResp.Enabled = True
                    Else
                        Bt_ImpFolhaResp.Enabled = False
                End If
                Exit Sub
        End If
        If .TextMatrix(.Row, 0) = "" Then
                Txt_Tipo.Enabled = False
                Meb_Dt.Enabled = False
                Bt_AplicarProva.Enabled = False
                Exit Sub
            Else
                Txt_Tipo.Enabled = True
                Meb_Dt.Enabled = True
                Bt_AplicarProva.Enabled = True
        End If
        '==============================================================================================
        'AVISA SE O ALUNO JA ESTA CURSANDO O LIMITE DE DISCIPLINAS
        'REVISAR ESTE CODIGO NAS PROXIMAS VERSOES
        
        'MaxDisciplCursando
        Dim RsMatrDisciplTMP    As Recordset
        Dim ii                  As Integer
        Dim iii                 As Integer
        Set RsMatrDisciplTMP = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND ISNULL(DtInicio)=FALSE AND ISNULL(DtConclusao) = TRUE")
        If RsMatrDisciplTMP.BOF And RsMatrDisciplTMP.EOF Then
                RsMatrDisciplTMP.Close
            Else
                
                RsMatrDisciplTMP.MoveLast
                
                If RsMatrDisciplTMP.RecordCount >= MaxDisciplCursando Then
                    RsMatrDisciplTMP.MoveFirst
                    iii = 0
                    For ii = 1 To RsMatrDisciplTMP.RecordCount
                        
                        If RsMatrDisciplTMP.Fields("DisciplinaID") = DisciplinaID Then
                                iii = 0
                                Exit For
                            Else
                                iii = 1
                        End If
                        RsMatrDisciplTMP.MoveNext
                    Next
                    If iii <> 0 Then
                        MsgBox "Matrícula cursando o número máximo (" & MaxDisciplCursando & ") de disciplinas.", vbInformation, "CESNet - Aviso"
                    End If
                    RsMatrDisciplTMP.Close
                    
                End If
        
        End If
        '==============================================================================================

        '==============================================================================================
        'CHECA SE HA NECESSIDA DE ORIENTAÇÃO E BLOQUEIA A SITUACAO SEM IMPERDIR A VISUALIZACAO
        If Orientacao = True Then
                Lb_Prova.ForeColor = vbRed
                Meb_Dt.ForeColor = vbRed
                Lb_Prova.Caption = .TextMatrix(.Row, 0) & " - " & .TextMatrix(.Row, 2)
                Bt_AplicarProva.Enabled = False
                Bt_ImpFolhaResp.Enabled = False
                Txt_Tipo.Enabled = False
                Meb_Dt.Enabled = False
                Txt_Tipo.Text = ""
                'Txt_Tipo.SetFocus
                Meb_Dt.Text = Date
                Exit Sub
        
        End If
        '============================================================================================
        If Trim(.TextMatrix(.Row, 3)) = "" Then
                Lb_Prova.ForeColor = &HFF0000
                Meb_Dt.ForeColor = &HFF0000
                Lb_Prova.Caption = .TextMatrix(.Row, 0) & " - " & .TextMatrix(.Row, 2)
                Bt_AplicarProva.Enabled = True
                If Trim(.TextMatrix(.Row, 3)) <> "" And Trim(.TextMatrix(.Row, 4)) = "" Then
                        Bt_ImpFolhaResp.Enabled = True
                    Else
                        Bt_ImpFolhaResp.Enabled = False
                End If
                Txt_Tipo.Enabled = True
                Meb_Dt.Enabled = True
                Txt_Tipo.Text = ""
                Txt_Tipo.SetFocus
                Meb_Dt.Text = Date
            Else
                Lb_Prova.ForeColor = &HFF&
                Meb_Dt.ForeColor = &HFF&
                Lb_Prova.Caption = .TextMatrix(.Row, 0) & " - " & .TextMatrix(.Row, 2)
                Meb_Dt.PromptInclude = False
                Meb_Dt.Text = .TextMatrix(.Row, 3)
                Meb_Dt.PromptInclude = True
                Txt_Tipo.Text = .TextMatrix(.Row, 4)
                Bt_AplicarProva.Enabled = False
                If Trim(.TextMatrix(.Row, 4)) <> "" And Trim(.TextMatrix(.Row, 5)) = "" Then
                        Bt_ImpFolhaResp.Enabled = True
                    Else
                        Bt_ImpFolhaResp.Enabled = False
                End If
                Txt_Tipo.Enabled = False
                Meb_Dt.Enabled = False
        End If

    End With
End Sub
Private Sub MstProvas()
    MSFG_Provas.Rows = 1
    MSFG_Provas.Rows = 2
    lin = 1
    Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND ISNULL(DtDevolucao)")
    ModuloID = RsEmprestimoModulo.Fields("ModuloID")
    Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND ModuloID = " & ModuloID)
    If RsTrafego.BOF And RsTrafego.EOF Then
        MsgBox "Modulo nao encontrado no Trafego. Por favor verifique.", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    RsTrafego.MoveFirst
    RefTrafegoID = RsTrafego.Fields("RefTrafegoID")
    Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RefTrafegoID & " ORDER BY NProva")
    If RsProvas.BOF And RsProvas.EOF Then
        MsgBox "Não existe nenhuma Prova cadastrada para essa Disciplina", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    RsProvas.MoveFirst
    Do Until RsProvas.EOF
        nProva = RsProvas.Fields("NProva")
        If IsNull(RsProvas.Fields("ModuloID")) Then
                        nmModulo = ""
                    Else
                        nmModulo = PgNomeModulo(RsProvas.Fields("ModuloID"))
                End If
        
        
        MSFG_Provas.TextMatrix(lin, 0) = nProva
        MSFG_Provas.TextMatrix(lin, 1) = nmModulo
        MSFG_Provas.TextMatrix(lin, 2) = RsProvas.Fields("Assunto")
        MSFG_Provas.TextMatrix(lin, 6) = RefTrafegoID
        Set RsMatriculaProvas = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "'")
        If RsMatriculaProvas.BOF And RsMatriculaProvas.EOF Then
            Else
                RsMatriculaProvas.MoveFirst
                If RsMatriculaProvas.Fields("Aprovado") = True Then
                        '*****REVER*****
                        'MASTRA OS DADOS DA PROVA APROVADA
                        MSFG_Provas.Rows = MSFG_Provas.Rows - 1
                        lin = lin - 1
                        'MSFG_Provas.TextMatrix(lin, 2) = IIf(IsNull(RsMatriculaProvas.Fields("DtAvaliacao")), " ", Format(RsMatriculaProvas.Fields("DtAvaliacao"), "00/##/####"))
                        'MSFG_Provas.TextMatrix(lin, 3) = IIf(IsNull(RsMatriculaProvas.Fields("Tipo")), " ", RsMatriculaProvas.Fields("Tipo"))
                        'MSFG_Provas.TextMatrix(lin, 4) = RsMatriculaProvas.Fields("Nota")
                    Else
                        If IsNull(RsMatriculaProvas.Fields("Nota")) Then
                            MSFG_Provas.TextMatrix(lin, 2) = IIf(IsNull(RsMatriculaProvas.Fields("DtAvaliacao")), " ", RsMatriculaProvas.Fields("DtAvaliacao"))
                            MSFG_Provas.TextMatrix(lin, 3) = IIf(IsNull(RsMatriculaProvas.Fields("Tipo")), " ", RsMatriculaProvas.Fields("Tipo"))
                        End If
                        '*MSFG_Provas.TextMatrix(lin, 2) = IIf(IsNull(RsMatriculaProvas.Fields("DtAvaliacao")), " ", RsMatriculaProvas.Fields("DtAvaliacao"))
                        '*MSFG_Provas.TextMatrix(lin, 3) = IIf(IsNull(RsMatriculaProvas.Fields("Tipo")), " ", RsMatriculaProvas.Fields("Tipo"))
                        'MSFG_Provas.TextMatrix(lin, 4) = RsMatriculaProvas.Fields("Nota")
                        Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "' ORDER BY Seq, NProva")
                        If RsProvasTMP.BOF And RsProvasTMP.EOF Then
                            Else
                                RsProvasTMP.MoveFirst
                                Do Until RsProvasTMP.EOF
                                    MSFG_Provas.Rows = MSFG_Provas.Rows + 1
                                    lin = lin + 1
                                    MSFG_Provas.TextMatrix(lin, 0) = RsProvasTMP.Fields("NProva")
                                    MSFG_Provas.TextMatrix(lin, 2) = RsProvas.Fields("Assunto")
                                    MSFG_Provas.TextMatrix(lin, 3) = RsProvasTMP.Fields("DtAvaliacao")
                                    MSFG_Provas.TextMatrix(lin, 4) = RsProvasTMP.Fields("Tipo")
                                    MSFG_Provas.TextMatrix(lin, 5) = IIf(SisNota = True, RsProvasTMP.Fields("Nota"), "NÃO")
                                    MSFG_Provas.TextMatrix(lin, 6) = RefTrafegoID
                                    MSFG_Provas.Row = lin
                                    MSFG_Provas.Col = 0
                                    MSFG_Provas.ColSel = MSFG_Provas.Cols - 1
                                    MSFG_Provas.FillStyle = flexFillRepeat
                                    MSFG_Provas.CellForeColor = &HFF&
                                    MSFG_Provas.Row = 0
                                    RsProvasTMP.MoveNext
                                Loop
                        End If
                End If
        End If
        RsProvas.MoveNext
        lin = lin + 1
        MSFG_Provas.Rows = MSFG_Provas.Rows + 1
    Loop
    MSFG_Provas.Rows = MSFG_Provas.Rows - 1
     'Alinha o Titulo das provas
    With MSFG_Provas
        '.Rows = MSFG_Provas.Rows - 1
        If .Rows = 1 Then
            .Rows = 2
        End If
        .Col = 2
        .ColSel = 2
        .Row = 1
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = 1
    End With
    'Organiza o grid pelo num. da prova
    With MSFG_Provas
        If .Rows = 1 Then
            .Rows = 2
        End If
        .Row = 1
        .RowSel = .Rows - 1
        .Col = 0
        .ColSel = 0 '.Cols - 1
        .FillStyle = flexFillRepeat
        .Sort = 1
    End With
End Sub

Private Sub Txt_Obs_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub Txt_Tipo_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 65 Or KeyAscii > 90 Then
        Beep
        KeyAscii = 0
    End If
End Sub
Private Sub MstTodasProvas()
    MSFG_Provas.Rows = 1
    MSFG_Provas.Rows = 2
    lin = 1
    Dim RsMatriculaSerie As Recordset
    Dim StatusProva As Boolean 'True - Aprovado False - Reprovado
    'Set RsEmprestimoModulo = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND ISNULL(DtDevolucao)")
    'ModuloID = RsEmprestimoModulo.Fields("ModuloID")
    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND Aprovado = False " & "ORDER BY SerieID")
    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
            MsgBox "Não existe nenhuma prova para esta Disciplina", vbInformation, "CESNet - Atenção"
            Exit Sub
        Else
            RsMatriculaSerie.MoveFirst
            
    End If
    'loop
    Do Until RsMatriculaSerie.EOF
        SerieID = RsMatriculaSerie.Fields("SerieID")
        '
        Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)  'ModuloID = " & ModuloID)
        If RsTrafego.BOF And RsTrafego.EOF Then
                MsgBox "Modulo nao encontrado no Trafego. Por favor verifique.", vbInformation, "CESNet - Aviso!"
                Exit Sub
            Else
                RsTrafego.MoveFirst
                'RefTrafegoID = RsTrafego.Fields("RefTrafegoID")
        End If
        Do Until RsTrafego.EOF
            RefTrafegoID = RsTrafego.Fields("RefTrafegoID")
            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RefTrafegoID & " ORDER BY NProva")
            If RsProvas.BOF And RsProvas.EOF Then
                MsgBox "Não existe nenhuma Prova cadastrada para essa Disciplina.", vbInformation, "CESNet - Aviso!"
                Exit Sub
            End If
            RsProvas.MoveFirst
            Do Until RsProvas.EOF
                
                nProva = RsProvas.Fields("NProva")
                'PEGA O NOME DO MODULO
                If IsNull(RsProvas.Fields("ModuloID")) Then
                        If IsNull(RsTrafego.Fields("ModuloID")) Or Trim(RsTrafego.Fields("ModuloID")) = "" Then
                                nmModulo = ""
                            Else
                                nmModulo = PgNomeModulo(RsTrafego.Fields("ModuloID"))
                        End If
                    Else
                        nmModulo = PgNomeModulo(RsProvas.Fields("ModuloID"))
                End If
                
                MSFG_Provas.TextMatrix(lin, 0) = nProva
                MSFG_Provas.TextMatrix(lin, 1) = nmModulo
                MSFG_Provas.TextMatrix(lin, 2) = IIf(IsNull(RsProvas.Fields("Assunto")), "", RsProvas.Fields("Assunto"))
                MSFG_Provas.TextMatrix(lin, 6) = RefTrafegoID
                'Set RsMatriculaProvas = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "'")
                Set RsMatriculaProvas = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND NProva = '" & nProva & "'")
                If RsMatriculaProvas.BOF And RsMatriculaProvas.EOF Then
                    Else
                       
                        RsMatriculaProvas.MoveFirst
                        StatusProva = RsMatriculaProvas.Fields("Aprovado")
                        If RsMatriculaProvas.Fields("Aprovado") = True Then
                            
                            MSFG_Provas.Rows = MSFG_Provas.Rows - 1
                            lin = lin - 1
                         End If
                            'Else
                                If IsNull(RsMatriculaProvas.Fields("Nota")) Then
                                    MSFG_Provas.TextMatrix(lin, 3) = IIf(IsNull(RsMatriculaProvas.Fields("DtAvaliacao")), " ", RsMatriculaProvas.Fields("DtAvaliacao"))
                                    MSFG_Provas.TextMatrix(lin, 4) = IIf(IsNull(RsMatriculaProvas.Fields("Tipo")), " ", RsMatriculaProvas.Fields("Tipo"))
                                End If
                                '*MSFG_Provas.TextMatrix(lin, 4) = RsMatriculaProvas.Fields("Nota")
                            If StatusProva = False Then
                                Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "' ORDER BY Seq, NProva")
                                'Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND NProva = '" & nProva & "' ORDER BY Seq, NProva")
                                If RsProvasTMP.BOF And RsProvasTMP.EOF Then
                                    Else
                                        
                                        RsProvasTMP.MoveFirst
                                        Do Until RsProvasTMP.EOF
                                            MSFG_Provas.Rows = MSFG_Provas.Rows + 1
                                            lin = lin + 1
                                            MSFG_Provas.TextMatrix(lin, 0) = RsProvasTMP.Fields("NProva")
                                            MSFG_Provas.TextMatrix(lin, 1) = nmModulo
                                            MSFG_Provas.TextMatrix(lin, 2) = RsProvas.Fields("Assunto")
                                            MSFG_Provas.TextMatrix(lin, 3) = RsProvasTMP.Fields("DtAvaliacao")
                                            MSFG_Provas.TextMatrix(lin, 4) = RsProvasTMP.Fields("Tipo")
                                            MSFG_Provas.TextMatrix(lin, 5) = IIf(SisNota = True, RsProvasTMP.Fields("Nota"), "NÃO")
                                            MSFG_Provas.TextMatrix(lin, 6) = RefTrafegoID
                                            Txt_Obs.Text = IIf(IsNull(RsProvasTMP.Fields("Obs")), "", RsProvasTMP.Fields("OBS"))
                                            MSFG_Provas.Row = lin
                                            MSFG_Provas.Col = 0
                                            MSFG_Provas.ColSel = MSFG_Provas.Cols - 1
                                            MSFG_Provas.FillStyle = flexFillRepeat
                                            MSFG_Provas.CellForeColor = &HFF&
                                            MSFG_Provas.Row = 0
                                            RsProvasTMP.MoveNext
                                        Loop
                                End If
                            End If
                End If
                RsProvas.MoveNext
                lin = lin + 1
                MSFG_Provas.Rows = MSFG_Provas.Rows + 1
            Loop
            RsTrafego.MoveNext
        Loop
        RsMatriculaSerie.MoveNext
    Loop
    MSFG_Provas.Rows = MSFG_Provas.Rows - 1
    'Alinha o Titulo das provas
    With MSFG_Provas
        '.Rows = MSFG_Provas.Rows - 1
        If .Rows = 1 Then
            .Rows = 2
        End If
        .Col = 2
        .ColSel = 2
        .Row = 1
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = 1
    End With
    'Organiza o grid pelo num. da prova
    With MSFG_Provas
        If .Rows = 1 Then
            .Rows = 2
        End If
        .Row = 1
        .RowSel = .Rows - 1
        .Col = 0
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .Sort = 1
    End With
End Sub
Private Function nsProva() As String
    On Error GoTo TrtErro
    Dim i As Integer
    Dim tmp As String
    For i = 1 To Len(Trim(MebMatricula.Text))
        If IsNumeric(Mid(Trim(MebMatricula.Text), i, 1)) Then
            tmp = tmp & Mid(Trim(MebMatricula.Text), i, 1)
        End If
    Next
    tmp = tmp & left("00", 2 - Len(Trim(PgIDEnsino(Lb_Ensino.Caption)))) & PgIDEnsino(Lb_Ensino.Caption) & _
           left("00", 2 - Len(Trim(PgIDDisciplina(Cb_Disciplina.Text)))) & PgIDDisciplina(Cb_Disciplina.Text) & _
           left(Lb_Prova.Caption, 3)
    
    For i = 1 To 16 Step 4
        nsProva = nsProva & " " & Mid(tmp, i, 4)
    Next
    nsProva = Trim(nsProva)
    Exit Function
TrtErro:
    MsgBox "Erro ao gerar numero de serie da prova", vbInformation, "CESNet - Aviso!"
    Call RegLogErros(Err.Number, "Erro ao gerar numero nsprova", "Avaliacao", UsuarioID)
    nsProva = "0000 0000 0000 0000"
End Function
Private Sub iFolhaResp()
    Dim Criterio As String
    Criterio = "SELECT * FROM MatriculaProva WHERE MatrID = '" & MebMatricula.Text & _
               "' AND EnsinoID = " & PgIDEnsino(Lb_Ensino.Caption) & _
               " AND DisciplinaID = " & PgIDDisciplina(Cb_Disciplina.Text) & _
               " AND NProva = '" & Trim(left(Lb_Prova.Caption, 3)) & "'"
               
    Call Relatorio(rptFolhaResposta, Criterio)
    
    rptFolhaResposta.Sections("Corpo").Controls.Item("lbData").Caption = Meb_Dt.Text
    rptFolhaResposta.Sections("Corpo").Controls.Item("lbNome").Caption = MebMatricula.Text & " - " & Lb_Nome.Caption
    
    rptFolhaResposta.Sections("Corpo").Controls.Item("lbCurso").Caption = Lb_Ensino.Caption
    rptFolhaResposta.Sections("Corpo").Controls.Item("lbDisciplina").Caption = Cb_Disciplina.Text
    rptFolhaResposta.Sections("Corpo").Controls.Item("lbTipo").Caption = Txt_Tipo.Text
    rptFolhaResposta.Sections("Corpo").Controls.Item("lbAvaliacao").Caption = Lb_Prova.Caption
    'rptFolhaResposta.Sections("Corpo").Controls.Item("lbNome").Caption = "NOME"
    rptFolhaResposta.Sections("Rodape").Controls.Item("lbLCNome").Caption = MebMatricula.Text & " - " & Lb_Nome.Caption
    
    rptFolhaResposta.PrintReport True
    BD_ADO.Close
    Exit Sub
    'If Form_Impressora.LoadFormCI(True, True, False, False, False, True, True, True, False, False) = False Then
    '    Exit Sub
    'End If
    'Call cPreview(4, UnidadeEnsinoNome, PgUA(UnidadeEnsino))
    'ObjPreview.Font = "Arial"
    'ObjPreview.FontBold = True
    'ObjPreview.FontItalic = False
    'ObjPreview.FontUnderline = False
    ''

    
    'ObjPreview.FontSize = 12
    'ObjPreview.Print
    'ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("FOLHA RESPOSTA") / 2)
    'ObjPreview.Print "FOLHA RESPOSTA"
    'ObjPreview.FontBold = CI.Negrito
    'ObjPreview.FontSize = CI.tFonte
    'ObjPreview.Print
    'ObjPreview.Print Tab(5); "Data: " & Meb_Dt.Text
    'ObjPreview.FontSize = 12
    'ObjPreview.FontBold = True
    'ObjPreview.Print Tab(5); "Matricula: " & MebMatricula.Text & " - " & Lb_Nome.Caption
    'ObjPreview.FontBold = False
    'ObjPreview.FontSize = 12
    'ObjPreview.Print
    'ObjPreview.Print Tab(5); "Ensino: " & Lb_Ensino.Caption; Tab(50); "Disciplina: " & Cb_Disciplina.Text
    'ObjPreview.Print Tab(5); "Avaliação: " & Lb_Prova.Caption
    'ObjPreview.Print Tab(5); "Tipo: " & Txt_Tipo.Text
    'ObjPreview.Print Tab(5); "Avaliador: ___________________________________"
    'ObjPreview.Print
    'ObjPreview.Print Tab(5); "Aluno: _______________________________________________"
    'ObjPreview.Print
    'ObjPreview.FontSize = 12
    'ObjPreview.Print Tab(5); "(   ) HABILITADO    (   ) NÃO HABILITADO"
    'ObjPreview.FontSize = 8
    'ObjPreview.Print
    'ObjPreview.Print Tab(5); "As respostas devem ser na ordem das questões usando caneta azul ou preta."
    'ObjPreview.CurrentY = ObjPreview.ScaleHeight - 1000
    'ObjPreview.CurrentX = ObjPreview.ScaleWidth - 5600
    'ObjPreview.Print "+-------------------------------------------------------------------------------------------------"
    'ObjPreview.CurrentX = ObjPreview.ScaleWidth - 5600
    'ObjPreview.FontBold = True
    'ObjPreview.Print "|                      V A L E    L A N C H E"
    'ObjPreview.FontBold = False
    'ObjPreview.CurrentX = ObjPreview.ScaleWidth - 5600
    'ObjPreview.Print "|   Data: " & Meb_Dt.Text
    'ObjPreview.CurrentX = ObjPreview.ScaleWidth - 5600
    'ObjPreview.Print "|   Matricula: " & MebMatricula.Text & " - " & Lb_Nome.Caption

    'Call ImpRodape
    'If CI.Preview = False Then
    '    ObjPreview.EndDoc
    'End If

    
End Sub
Public Sub ReceberInformacoes(Matricula As String, Disciplina As String)
    MebMatricula.Text = Matricula
    Cb_Disciplina.Clear
    Cb_Disciplina.AddItem Disciplina
    Cb_Disciplina.Text = Cb_Disciplina.List(0)
    Me.Show
End Sub
