VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Notas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Notas"
   ClientHeight    =   6315
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   11130
   Icon            =   "Form_Notas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6315
   ScaleWidth      =   11130
   Begin VB.CommandButton btFoto 
      Caption         =   "&Foto"
      Height          =   735
      Left            =   9540
      Picture         =   "Form_Notas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   24
      Top             =   3600
      Width           =   1515
   End
   Begin VB.Frame Frame_Sintetico 
      Caption         =   "Notas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   180
      TabIndex        =   16
      Top             =   4500
      Width           =   10905
      Begin VB.TextBox Txt_ObsS 
         Height          =   555
         Left            =   945
         MaxLength       =   200
         MultiLine       =   -1  'True
         TabIndex        =   23
         Top             =   990
         Width           =   6315
      End
      Begin VB.ComboBox Cb_Aprovado 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form_Notas.frx":0614
         Left            =   945
         List            =   "Form_Notas.frx":0621
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   585
         Width           =   1500
      End
      Begin VB.CommandButton Bt_GrvNotaS 
         Caption         =   "Gravar Nota"
         Enabled         =   0   'False
         Height          =   960
         Left            =   7560
         Picture         =   "Form_Notas.frx":0632
         Style           =   1  'Graphical
         TabIndex        =   17
         Top             =   630
         Width           =   3045
      End
      Begin MSMask.MaskEdBox Meb_Nota 
         Height          =   315
         Left            =   915
         TabIndex        =   25
         Top             =   540
         Width           =   675
         _ExtentX        =   1191
         _ExtentY        =   556
         _Version        =   393216
         PromptInclude   =   0   'False
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Obs.:"
         Height          =   195
         Left            =   495
         TabIndex        =   22
         Top             =   990
         Width           =   420
      End
      Begin VB.Label lbSisNota 
         Alignment       =   1  'Right Justify
         Caption         =   "Aprovado:"
         Height          =   195
         Left            =   180
         TabIndex        =   21
         Top             =   645
         Width           =   735
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Prova:"
         Height          =   165
         Left            =   420
         TabIndex        =   20
         Top             =   225
         Width           =   495
      End
      Begin VB.Label Lb_Prova 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   945
         TabIndex        =   19
         Top             =   180
         Width           =   9795
      End
   End
   Begin VB.Frame Frame_ExclProva 
      Caption         =   "Exclusão de Prova:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1695
      Left            =   180
      TabIndex        =   12
      Top             =   4560
      Visible         =   0   'False
      Width           =   10905
      Begin VB.CommandButton Bt_ExclProva 
         Caption         =   "Excluir Prova"
         Enabled         =   0   'False
         Height          =   960
         Left            =   7695
         Picture         =   "Form_Notas.frx":093C
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   540
         Width           =   3045
      End
      Begin VB.Label Lb_ProvaE 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   630
         TabIndex        =   14
         Top             =   240
         Width           =   9795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Prova:"
         Height          =   165
         Left            =   90
         TabIndex        =   13
         Top             =   285
         Width           =   495
      End
   End
   Begin VB.Frame Frame2 
      Height          =   510
      Left            =   7200
      TabIndex        =   9
      Top             =   315
      Width           =   3840
      Begin VB.OptionButton Opt_Funcao 
         Caption         =   "Excluir Prova"
         Height          =   240
         Index           =   1
         Left            =   1935
         TabIndex        =   11
         Top             =   180
         Width           =   1500
      End
      Begin VB.OptionButton Opt_Funcao 
         Caption         =   "Lançar Nota"
         Height          =   240
         Index           =   0
         Left            =   135
         TabIndex        =   10
         Top             =   180
         Value           =   -1  'True
         Width           =   1185
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2610
      Left            =   75
      TabIndex        =   1
      Top             =   855
      Width           =   10980
      Begin MSFlexGridLib.MSFlexGrid MSFG_Provas 
         Height          =   2370
         Left            =   90
         TabIndex        =   8
         Top             =   180
         Width           =   10815
         _ExtentX        =   19076
         _ExtentY        =   4180
         _Version        =   393216
         Cols            =   9
         FixedCols       =   0
         Enabled         =   0   'False
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_Notas.frx":0C46
      End
   End
   Begin MSMask.MaskEdBox MebMatricula 
      Height          =   375
      Left            =   975
      TabIndex        =   3
      Top             =   3600
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
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
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   195
      Left            =   420
      TabIndex        =   7
      Top             =   4125
      Width           =   510
   End
   Begin VB.Label Lb_Aluno 
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
      Left            =   975
      TabIndex        =   6
      Top             =   4065
      Width           =   7920
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Professor:"
      Height          =   195
      Left            =   60
      TabIndex        =   5
      Top             =   450
      Width           =   735
   End
   Begin VB.Label Lb_Professor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   900
      TabIndex        =   4
      Top             =   420
      Width           =   6120
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Matricula:"
      Height          =   195
      Left            =   135
      TabIndex        =   2
      Top             =   3735
      Width           =   795
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "LANÇAR NOTAS"
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
      Width           =   11205
   End
End
Attribute VB_Name = "Form_Notas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsMatricula             As Recordset
Dim RsTrafego               As Recordset
Dim RsDisciplina            As Recordset
Dim RsMatriculaEnsino       As Recordset
Dim RsMatriculaDisciplina   As Recordset
Dim RsMatriculaSerie        As Recordset
Dim RsMatriculaProva        As Recordset
Dim RsProva                 As Recordset
Dim RsSeries                As Recordset
Dim RsProfessor             As Recordset
Dim RsProfessorDisciplina   As Recordset
Dim RsProvasTMP             As Recordset
Dim RsUnidade               As Recordset
Dim RsOrientacao            As Recordset

Dim MatrID                  As String
Dim EnsinoID                As Integer
Dim DisciplinaID            As Integer

Dim SerieID                 As Integer
Dim RefTrafegoID            As Integer
Dim nProva                  As String
Dim Assunto                 As String
Dim DtAvaliacao             As Date
Dim UnidEns                 As String
'Dim Aprovado as Boolean
Dim ProfID                  As Integer
Dim tmp                     As Integer
Dim lin                     As Integer

Dim Funcao As Boolean 'True - Lancar notas // False - Ecluir notas



Private Sub GrvProva(op As String)
    'S = Sintetica (Media para aprovacao)
    'A = Analitica (Meb_nota.text)
    
    
    Dim Aprovado As String 'Colocar SIM ou NAO para Aprovado
    
    
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    
    Select Case op
        Case "S"
            If Trim(Cb_Aprovado.Text) = "" Then
                MsgBox "Favor informar se o aluno(a) foi aprovado(a)!", vbInformation, "CESNet - Aviso"
                Exit Sub
            End If
            Aprovado = Cb_Aprovado.Text
        Case "A"
            If Trim(Meb_Nota.Text) = "" Then
                MsgBox "Favor informar a nota do aluno(a)!", vbInformation, "CESNet - Aviso"
                Exit Sub
            End If
            If Meb_Nota.Text >= NotaMedia Then
                    Aprovado = "SIM"
                Else
                    Aprovado = "NAO"
            End If
    End Select
    
    
    MSFG_Provas.Enabled = False
    RefTrafegoID = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0)
    EnsinoID = PgIDEnsino(MSFG_Provas.TextMatrix(MSFG_Provas.Row, 1))
    nProva = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 5)
    DtAvaliacao = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 4)
    
    
    'CHECA A PROVA
    Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "'")
    If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
        MsgBox "Erro ao procurar pela prova chame o Suporte.", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    RsMatriculaProva.MoveFirst
    '*********** APROVADO
    'If Cb_Aprovado.Text = "SIM" Then
    If Aprovado = "SIM" Then
            RsMatriculaProva.Edit
            RsMatriculaProva.Fields("Nota") = IIf(op = "S", NotaMedia, Meb_Nota.Text)
            RsMatriculaProva.Fields("Aprovado") = True
            RsMatriculaProva.Fields("Status") = "HB"
            RsMatriculaProva.Fields("Obs") = IIf(Trim(Txt_ObsS.Text) = "", Null, Trim(Txt_ObsS.Text))
            RsMatriculaProva.Fields("ProfIDN") = ProfID
            RsMatriculaProva.Fields("UsuarioIDN") = UsuarioID
            RsMatriculaProva.Fields("DtHrN") = Now
            RsMatriculaProva.Update
            '************************************************************************************
            'APAGA AS PROVAS TEMPORARIAS
            'Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MatrId = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "'")
            'If RsProvasTMP.BOF And RsProvasTMP.EOF Then
            '    Else
            '        RsProvasTMP.MoveFirst
            '        Do Until RsProvasTMP.EOF
            '            RsProvasTMP.Delete
            '            RsProvasTMP.MoveNext
            '        Loop
            'End If
            '************************************************************************************
            
            'Checar se o aluno precisa de orientacao
            If NumRepro = "0" Then
                Call ChkProvas(RsMatriculaProva.Fields("RefTrafegoID"))
            End If
            
            
            'CHECAR SE CONCLUIU A Disciplina
            Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE RefTrafegoID = " & RefTrafegoID)
            EnsinoID = RsTrafego.Fields("EnsinoID")
            DisciplinaID = RsTrafego.Fields("DisciplinaID")
            SerieID = RsTrafego.Fields("SerieID")
            
                    If ChecarProvas(MatrID, EnsinoID, DisciplinaID, SerieID) = True Then
                            GrvMatriculaSerie
                        End If
                    
                    If Chk_ConcDisciplina(MatrID, EnsinoID, DisciplinaID) = True Then
                        GrvMatriculaDisciplina
                        '*****************************
                        'Checar se nao existe mais Disciplinas
                        If Chk_ConcEnsino(MatrID, EnsinoID) = True Then
                            Call Grv_ConcEnsino(MatrID, EnsinoID)
                            
                        End If
                    End If
        Else
            '********* NAO APROVADO
            RsMatriculaProva.Edit
            RsMatriculaProva.Fields("Nota") = IIf(op = "S", "0", Trim(Meb_Nota.Text))
             RsMatriculaProva.Fields("Status") = "NH"
            RsMatriculaProva.Update
            Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP")
            RsProvasTMP.AddNew
            RsProvasTMP.Fields("MPID") = RsMatriculaProva.Fields("ID")
            RsProvasTMP.Fields("MatrID") = MebMatricula.Text
            RsProvasTMP.Fields("RefTrafegoID") = RsMatriculaProva.Fields("RefTrafegoID")
            RsProvasTMP.Fields("EnsinoID") = RsMatriculaProva.Fields("EnsinoID")
            RsProvasTMP.Fields("DisciplinaID") = RsMatriculaProva.Fields("DisciplinaID")
            RsProvasTMP.Fields("NProva") = RsMatriculaProva.Fields("NProva")
            RsProvasTMP.Fields("Nota") = IIf(op = "S", NotaMedia, Meb_Nota.Text)
            RsProvasTMP.Fields("Tipo") = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 7)
            RsProvasTMP.Fields("DtAvaliacao") = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 4)
            RsProvasTMP.Fields("Obs") = IIf(Trim(Txt_ObsS.Text) = "", Null, Trim(Txt_ObsS.Text))
            RsProvasTMP.Fields("DtHrAv") = RsMatriculaProva.Fields("DtHrAv")
            RsProvasTMP.Fields("UsuarioIDAv") = RsMatriculaProva.Fields("UsuarioIDAv")
            RsProvasTMP.Fields("ProfIDN") = ProfID
            RsProvasTMP.Fields("UsuarioIDN") = UsuarioID
            RsProvasTMP.Fields("DtHrN") = Now
            RsProvasTMP.Update
            'Checar se o aluno precisa de orientacao
            Call ChkProvas(RsMatriculaProva.Fields("RefTrafegoID"))
    End If
    
    
    RegLog MebMatricula.Text, "LANCAMENTO NOTA: " & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 1) & "/" & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 2) & "/" & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 3) & "/" & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 4) & "/" & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 5) & "/" & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 6) & "/" & _
                               Aprovado & "/" & _
                               Lb_Professor.Caption & " (" & ProfID & ")" & "/" & _
                               UsuarioID
                               'MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0) & "/" & _
                               MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0)
    
    MstProvas
    Cb_Aprovado.Text = " "
    Meb_Nota.Text = ""
    Cb_Aprovado.Enabled = False
    Bt_GrvNotaS.Enabled = False
    Lb_Prova.Caption = ""
    Lb_Prova.Caption = ""
    Txt_ObsS.Text = ""

End Sub


Private Sub Bt_ExclProva_Click()
    If ChkAcesso(Me.Name, "E") = False Then Exit Sub
    MSFG_Provas.Enabled = False
    RefTrafegoID = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0)
    nProva = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 5)
    DtAvaliacao = MSFG_Provas.TextMatrix(MSFG_Provas.Row, 4)
    DisciplinaID = PgIDDisciplina(MSFG_Provas.TextMatrix(MSFG_Provas.Row, 2))
    
    If MsgBox("Deseja realmente excluir esta prova n. " & nProva & " - " & MSFG_Provas.TextMatrix(MSFG_Provas.Row, 6) & "?", vbYesNo + vbInformation, "CESNet - Aviso!") = vbYes Then
        'Autentica o usuario
        If Form_AutenticacaoUsuario.CarregarForm = False Then
            MSFG_Provas.Enabled = True
            Bt_ExclProva.Enabled = False
            Exit Sub
        End If
        
        'BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID = '" & MatrID & _
                   "' AND RefTrafegoID = " & RefTrafegoID & _
                   " AND NProva = '" & nProva & "'"
        BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID = '" & MatrID & _
                   "' AND EnsinoID = " & EnsinoID & _
                   " AND DisciplinaID = " & DisciplinaID & _
                   " AND NProva = '" & nProva & "'"
        SerieID = PegSerie(RefTrafegoID)
        Call RegLog(MatrID, "Prova Excluida - Ensino " & PgNomeEnsino(EnsinoID) & _
                    ", Disciplina " & PgNomeDisciplina(DisciplinaID) & _
                    ",  Prova " & nProva & _
                    ", avaliado em " & DtAvaliacao & _
                    ", tipo " & MSFG_Provas.TextMatrix(MSFG_Provas.Row, MSFG_Provas.Cols - 2) & _
                    ", nota " & MSFG_Provas.TextMatrix(MSFG_Provas.Row, MSFG_Provas.Cols - 1) & "/Apto")
        
        'Checar Serie
        Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & _
                        "' AND EnsinoID = " & EnsinoID & _
                        " AND DisciplinaID = " & DisciplinaID & _
                        " AND SerieID = " & SerieID)
        If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                MsgBox "Erro ao localizar série. Favor avise ao suporte.", vbInformation, "CESNet - Aviso"
            Else
                RsMatriculaSerie.MoveFirst
                RsMatriculaSerie.Edit
                RsMatriculaSerie.Fields("DtFinal") = Null
                RsMatriculaSerie.Fields("Aprovado") = False
                RsMatriculaSerie.Update
                RsMatriculaSerie.Close
        End If
        'Checar Disciplina
        Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE " & _
                                    "MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & _
                                    " AND DisciplinaID = " & DisciplinaID)
        If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                MsgBox "Erro ao localizar disciplina. Favor avise ao suporte.", vbInformation, "CESNet - Aviso"
            Else
                 RsMatriculaDisciplina.MoveFirst
                 If IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) Then
                    Else
                        RsMatriculaDisciplina.Edit
                        RsMatriculaDisciplina.Fields("DtConclusao") = Null
                        RsMatriculaDisciplina.Fields("Cidade") = Null
                        RsMatriculaDisciplina.Fields("Local") = Null
                        RsMatriculaDisciplina.Fields("UF") = Null
                        RsMatriculaDisciplina.Update
                        RsMatriculaDisciplina.Close
                 End If
        End If
        'Checar Ensino
        Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE " & _
                                    "MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
        If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                MsgBox "Erro ao localizar disciplina. Favor avise ao suporte.", vbInformation, "CESNet - Aviso"
            Else
                 RsMatriculaEnsino.MoveFirst
                 If IsNull(RsMatriculaEnsino.Fields("DtFinal")) Then
                    Else
                        RsMatriculaEnsino.Edit
                        RsMatriculaEnsino.Fields("DtFinal") = Null
                        RsMatriculaEnsino.Fields("Local") = Null
                        RsMatriculaEnsino.Update
                        RsMatriculaEnsino.Close
                 End If
        End If
        LstProvas
    End If
    RefTrafegoID = Empty
    nProva = Empty
    DtAvaliacao = Empty
    DisciplinaID = Empty

    Lb_ProvaE.Caption = ""
    Bt_ExclProva.Enabled = False
    MSFG_Provas.Enabled = True
End Sub

Private Sub Bt_GrvNotaS_Click()
    GrvProva (IIf(SisNota = True, "A", "S"))
    
End Sub



Private Sub GrvMatriculaSerie()
    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
            MsgBox "Erro no acesso ao Banco de Dados. Por favor chame o Suporte", vbInformation, "CESNet - Aviso!"
        Else
            RsMatriculaProva.MoveFirst
            RsMatriculaSerie.Edit
            RsMatriculaSerie.Fields("DtIni") = RsMatriculaProva.Fields("DtAvaliacao")
            RsMatriculaProva.MoveLast
            RsMatriculaSerie.Fields("DtFinal") = RsMatriculaProva.Fields("DtAvaliacao")
            RsMatriculaSerie.Fields("Aprovado") = True
            RsMatriculaSerie.Update
    End If
End Sub

Private Sub GrvMatriculaDisciplina()
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            MsgBox "Erro no acesso ao Banco de Dados. Por favor chame o Suporte", vbInformation, "CESNet - Aviso!"
        Else
            'Set RsUnidade = BD.OpenRecordset("SELECT * FROM Unidades WHERE UnidID = '" & UnidadeEnsino & "'")
            'If RsUnidade.BOF And RsUnidade.EOF Then
            '        MsgBox "Erro na buscada Unidade de Ensino. Verifique o CADASTRO DE UNIDADES E CHAME O SUPORTE", vbInformation, "CESNet - Aviso!"
            '        UnidEns = "NESTA UNIDADE DE ENSINO"
            '    Else
            '        RsUnidade.MoveFirst
            '        UnidEns = RsUnidade.Fields("Nome")
            'End If
            
            '=============================================================================================
            'FUNCAO TEMPORARIA PARA PEGAR O INICIO DA DISCIPLINA
            Dim RsMatrSerieTMP As Recordset
            Dim DtInicio As String
            
            Set RsMatrSerieTMP = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY SerieID")
            If RsMatrSerieTMP.BOF And RsMatrSerieTMP.EOF Then
                    RsMatrSerieTMP.Close
                    DtInicio = DtAvaliacao
                Else
                    RsMatrSerieTMP.MoveFirst
                    ''Debug.Print RsMatrSerieTMP.Fields("SerieID")
                    DtInicio = RsMatrSerieTMP.Fields("DtIni")
                    RsMatrSerieTMP.Close
            End If
            '================================================================================================
            RsMatriculaDisciplina.MoveFirst
            
            RsMatriculaDisciplina.Edit
            RsMatriculaDisciplina.Fields("DtInicio") = DtInicio
            RsMatriculaDisciplina.Fields("DtConclusao") = DtAvaliacao
            RsMatriculaDisciplina.Fields("Local") = UnidadeEnsinoNome 'UnidEns
            RsMatriculaDisciplina.Fields("Cidade") = PgDadosUnid(UnidadeEnsino).Municipio
            RsMatriculaDisciplina.Fields("UF") = PgDadosUnid(UnidadeEnsino).UF
            RsMatriculaDisciplina.Update
    End If
End Sub

Private Sub btFoto_Click()
    Form_ExibirImagem.ExibirFoto (MatrID)
End Sub

Private Sub Cb_Aprovado_Click()
'    If Cb_Aprovado.Text = "SIM" Then
'            Txt_ObsS.Enabled = False
'        ElseIf Cb_Aprovado.Text = "NÃO" Then
'            Txt_ObsS.Enabled = True
'        Else
'            Txt_ObsS.Enabled = False
'
'    End If
End Sub


Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
    'Acrescentar para quando o form tiver o foco acertar os dados
    SistLanca
End Sub
Private Sub SistLanca()
    Funcao = True
    Frame_ExclProva.Visible = False
    If SisNota = True Then
            lbSisNota.Caption = "Nota:"
            Cb_Aprovado.Visible = False
            Meb_Nota.Visible = True
            'Frame_Analitico.Visible = True
            'Frame_Sintetico.Visible = False
        Else
            lbSisNota.Caption = "Aprovado:"
            Cb_Aprovado.Visible = True
            Meb_Nota.Visible = False
            'Frame_Analitico.Visible = False
            'Frame_Sintetico.Visible = True
    End If
End Sub
Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores WHERE UsuarioID = " & UsuarioID)
    If RsProfessor.BOF And RsProfessor.EOF Then
            MsgBox "Usuário não cadastrado para esta sessão.", vbInformation, "CESNet - Aviso!"
            tmp = 1
            ProfID = 0
            RsProfessor.Close
            Exit Sub
            
         
            
            
        Else
            RsProfessor.MoveFirst
            ProfID = RsProfessor.Fields("ProfID")
            Lb_Professor.Caption = RsProfessor.Fields("Nome")
            'Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina WHERE Chv = '" & Usuario & "'")
    End If
    'If SisNota = True Then
    '        Frame_Analitico.Visible = True
    '        Frame_Sintetico.Visible = False
    '    Else
    '        Frame_Analitico.Visible = False
    '        Frame_Sintetico.Visible = True
    'End If
    'Me.Top = 0
    'Me.Left = 0
End Sub





Private Sub Meb_Nota_Change()
    If Val(Meb_Nota.Text) >= NotaMedia Then
            Txt_ObsS.Enabled = False
        Else
            Txt_ObsS.Enabled = True
    End If
End Sub

Private Sub Meb_Nota_GotFocus()
    Meb_Nota.SelStart = 0
    Meb_Nota.SelLength = 4
End Sub


Private Sub Meb_Nota_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Bt_GrvNotaS.SetFocus
    End If
End Sub

Private Sub Frame_Analitico_DragDrop(Source As Control, x As Single, y As Single)

End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub

Private Sub MebMatricula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
        MatrID = formBuscar.IniciarBusca("Matriculas")
        If MatrID = 0 Then Exit Sub
        MebMatricula.Text = MatrID
        MebMatricula_KeyPress (13)
    End If
End Sub

Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        MatrID = MebMatricula.Text
        Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID = '" & MatrID & "'")
        If RsMatricula.BOF And RsMatricula.EOF Then
            MsgBox "Matricula Invalida!", vbInformation, "CESNet - Aviso!"
            Lb_ProvaE.Caption = ""
            'lb_Prova.Caption = ""
            Lb_Prova.Caption = ""
            MebMatricula.SetFocus
            MebMatricula_GotFocus
            Exit Sub
        End If
        RsMatricula.MoveFirst
        Lb_Aluno.Caption = RsMatricula.Fields("Nome")
        If Funcao = True Then
                MstProvas
            Else
                LstProvas
        End If
    End If
End Sub

Private Sub MSFG_Provas_Click()
    With MSFG_Provas
        If .TextMatrix(.Row, 2) = "" Or .TextMatrix(.Row, 3) = "" Then
            Exit Sub
        End If
        DisciplinaID = PgIDDisciplina(.TextMatrix(.Row, 2))
        
        If PgProfDisciplina(ProfID, DisciplinaID) = True Then
                Lb_ProvaE.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Bt_ExclProva.Enabled = True
                        
                'Analitico
                Lb_ProvaE.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Meb_Nota.Enabled = True
                'Bt_GrvNotaS.Enabled = True
                'Sistetico
                Lb_Prova.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Cb_Aprovado.Enabled = True
                Bt_GrvNotaS.Enabled = True
                'Exclusao de prova
                Lb_ProvaE.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Bt_ExclProva.Enabled = True

            Else
                        
                'Analitico
                Lb_ProvaE.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Meb_Nota.Enabled = False
                Bt_GrvNotaS.Enabled = False
                'Sistetico
                Lb_Prova.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Cb_Aprovado.Enabled = False
                Bt_GrvNotaS.Enabled = False
                'Exclusao de prova
                Lb_ProvaE.Caption = .TextMatrix(.Row, 5) & " - " & .TextMatrix(.Row, 6)
                Bt_ExclProva.Enabled = False
                MsgBox "Usuário não autorizado para alterar esta Disciplina!", vbInformation, "CESNet - Aviso!"
        End If
    End With
End Sub
Private Sub MstProvas() 'Mostra as provas que falta lançar nota
    On Error GoTo TrtErroMst
    MSFG_Provas.Rows = 1
    MSFG_Provas.Rows = 2
    Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND IsNull(Nota) ") ' AND Aprovado = False")
    If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
            'MsgBox "Nenhuma Prova para esta matricula!", vbInformation, "CESNet - Aviso!"
            MebMatricula.PromptInclude = False
            MebMatricula.Text = ""
            Lb_Aluno.Caption = ""
            Lb_Prova.Caption = ""
            Txt_ObsS.Text = ""
            MebMatricula.PromptInclude = True
            MebMatricula.SetFocus
            MebMatricula_GotFocus
            Exit Sub
        Else
            RsMatriculaProva.MoveFirst
            lin = 1
            'RefTrafegoID = RsMatriculaProva.Fields("RefTrafegoID")
            Do Until RsMatriculaProva.EOF
                RefTrafegoID = RsMatriculaProva.Fields("RefTrafegoID")
                Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE RefTrafegoID = " & RefTrafegoID)
                If RsTrafego.BOF And RsTrafego.EOF Then
                        'Erro ao localizar o trafego exclui a prova aplicada
                        MsgBox "Erro na busca da referencia do Trafego, tal erro deve ter sido causado devido a mudanças na grade de provas. " & Chr(13) & _
                               "A prova abaixo deverá será exluida e o usuario(a) deve adaptar a prova as novas condições impostas pela nova grade." & Chr(13) & _
                               "Ensino: " & IIf(PgNomeEnsino(IIf(IsNull(RsMatriculaProva.Fields("EnsinoID")), 0, RsMatriculaProva.Fields("EnsinoID"))) = 0, "<< Nenhum >>", PgNomeEnsino(IIf(IsNull(RsMatriculaProva.Fields("EnsinoID")), 0, RsMatriculaProva.Fields("EnsinoID")))) & Chr(13) & _
                               "Disciplina: " & IIf(PgNomeDisciplina(IIf(IsNull(RsMatriculaProva.Fields("DisciplinaID")), 0, RsMatriculaProva.Fields("DisciplinaID"))) = 0, "<< Nenhum >>", PgNomeDisciplina(IIf(IsNull(RsMatriculaProva.Fields("DisciplinaID")), 0, RsMatriculaProva.Fields("DisciplinaID")))) & Chr(13) & _
                               "Prova: " & RsMatriculaProva.Fields("NProva") & " - " & RsMatriculaProva.Fields("Assunto") & Chr(13) _
                               , vbInformation, "CESNet - Erro no Trafego"
                        If MsgBox("Deseja realmente excluir a Prova n. " & RsMatriculaProva.Fields("NProva") & "?", vbInformation + vbYesNo, "CESNet - Exclusao de prova") = vbYes Then
                                
                                Call RegLog(MatrID, "Prova excluida devido a erro no trafego. - Prova: " & RsMatriculaProva.Fields("NProva") & _
                                            " Avaliacao: " & RsMatriculaProva.Fields("DtAvaliacao") & _
                                            " Tipo: " & RsMatriculaProva.Fields("Tipo") & _
                                            " Nota: " & IIf(IsNull(RsMatriculaProva.Fields("Nota")), "(Nao corrigido)", RsMatriculaProva.Fields("Nota")) & _
                                            "/" & IIf(RsMatriculaProva.Fields("Aprovado") = True, "Apto", "inapto"))
                                           
                                
                                BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID
                                MsgBox "Prova Excluida com sucesso", vbInformation, "CESNet - Aviso"
                            Else
                                'Exit Sub
                        End If
                        Exit Sub
                        
                    Else
                        RsTrafego.MoveFirst
                        DisciplinaID = RsTrafego.Fields("DisciplinaID")
                End If
                Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina WHERE ID = " & DisciplinaID)
                If RsDisciplina.BOF And RsDisciplina.EOF Then
                        MsgBox "Erro na Referencia da Disciplina chame o Suporte.", vbInformation, "CESNet - Aviso!"
                        Exit Sub
                    Else
                        RsDisciplina.MoveFirst
                End If
                nProva = RsMatriculaProva.Fields("NProva")
                Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "'")
                If RsProva.BOF And RsProva.EOF Then
                        'MsgBox "Erro na Referencia da Prova chame o Suporte.", vbInformation, "CESNet - Aviso!"
                        'Exit Sub
                        Assunto = IIf(IsNull(RsMatriculaProva.Fields("Assunto")), " ", RsMatriculaProva.Fields("Assunto"))
                    Else
                        RsProva.MoveFirst
                        Assunto = IIf(IsNull(RsProva.Fields("Assunto")), " ", RsProva.Fields("Assunto"))
                End If
                MSFG_Provas.TextMatrix(lin, 0) = RefTrafegoID
                MSFG_Provas.TextMatrix(lin, 1) = PgNomeEnsino(RsProva.Fields("EnsinoID"))
                MSFG_Provas.TextMatrix(lin, 2) = RsDisciplina.Fields("Descr")
                MSFG_Provas.TextMatrix(lin, 3) = PgNomeModulo(RsProva.Fields("ModuloID"))
                MSFG_Provas.TextMatrix(lin, 4) = IIf(IsNull(RsMatriculaProva.Fields("DtAvaliacao")), " ", RsMatriculaProva.Fields("DtAvaliacao"))
                MSFG_Provas.TextMatrix(lin, 5) = nProva
                MSFG_Provas.TextMatrix(lin, 6) = Assunto
                MSFG_Provas.TextMatrix(lin, 7) = RsMatriculaProva.Fields("Tipo")
                MSFG_Provas.TextMatrix(lin, 8) = IIf(IsNull(RsMatriculaProva.Fields("Nota")), " ", RsMatriculaProva.Fields("Nota"))
                RsMatriculaProva.MoveNext
                lin = lin + 1
                MSFG_Provas.Rows = MSFG_Provas.Rows + 1
            Loop
            MSFG_Provas.Rows = MSFG_Provas.Rows - 1
            'Alinha o Titulo das provas
            With MSFG_Provas
                
                If .Rows = 1 Then
                    .Rows = 2
                End If
                .Col = 5
                .ColSel = 5
                .Row = 1
                .RowSel = .Rows - 1
                .FillStyle = flexFillRepeat
                .CellAlignment = 1
                
                .Row = 1
            End With
    End If
    MSFG_Provas.Enabled = True
    Exit Sub
TrtErroMst:
    Resume Next
End Sub
Private Sub ChkProvas(RefTrafego As Integer)
    Dim sSQL As String
    'Checa se o aluno precisa de Orientacao

    sSQL = "SELECT * FROM ProvasTMP WHERE " & _
         "MatrID = '" & MatrID & "' AND " & _
         "EnsinoID = " & EnsinoID & " AND " & _
         "DisciplinaID = " & DisciplinaID & " AND " & _
         "NProva = '" & nProva & "'"
         

    Set RsProvasTMP = BD.OpenRecordset(sSQL)
    If RsProvasTMP.BOF And RsProvasTMP.EOF Then
        
        Else
            RsProvasTMP.MoveLast
            If RsProvasTMP.RecordCount >= NumRepro Then
                Call GrvOrientacao(RsProvasTMP.Fields("EnsinoID"), RsProvasTMP.Fields("DisciplinaID"))
            End If
    End If
'
'    Set RsProvasTMP = BD.OpenRecordset("SELECT Trafego.RefTrafegoID, Trafego.EnsinoID, Trafego.DisciplinaID, ProvasTMP.MatrID, ProvasTMP.NProva " & _
'                                        "FROM Trafego INNER JOIN ProvasTMP ON Trafego.RefTrafegoID = ProvasTMP.RefTrafegoID " & _
'                                        "WHERE (((ProvasTMP.MatrID)='" & MatrID & "') AND ((Trafego.RefTrafegoID)=" & RefTrafego & "))") ' AND ((Trafego.DisciplinaID)=" & DisciplinaID & "))")
'

'
'    If RsProvasTMP.BOF And RsProvasTMP.EOF Then
'            'MsgBox "NAO Acho"
'            If NumRepro = 0 Then
'                Set RsProvasTMP = BD.OpenRecordset("SELECT * FROM Trafego WHERE RefTrafegoID = " & RefTrafego)
'                If RsProvasTMP.BOF And RsProvasTMP.EOF Then
'                        MsgBox "Erro ao loc. Ensino e Disciplina na tab. trafego na função ChkProvas. Registro cancelado.", vbInformation, "CESNet - Aviso"
'                        Exit Sub
'                    Else
'                        RsProvasTMP.MoveFirst
'                        Call GrvOrientacao(RsProvasTMP.Fields("EnsinoID"), RsProvasTMP.Fields("DisciplinaID"))
'                End If
'            End If
'        Else
'            RsProvasTMP.MoveLast
'            If RsProvasTMP.RecordCount >= NumRepro Then
'                Call GrvOrientacao(RsProvasTMP.Fields("EnsinoID"), RsProvasTMP.Fields("DisciplinaID"))
'            End If
'    End If


End Sub
Private Sub GrvOrientacao(ID_Ensino As Integer, ID_Disciplina As Integer)
    Set RsOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao") ' WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
    RsOrientacao.AddNew
    RsOrientacao.Fields("MatrID") = MatrID
    RsOrientacao.Fields("EnsinoID") = ID_Ensino
    RsOrientacao.Fields("DisciplinaID") = ID_Disciplina
    RsOrientacao.Fields("DtAvaliacao") = DtAvaliacao
    RsOrientacao.Fields("DtHr") = Now
    RsOrientacao.Update
    RsOrientacao.Close
End Sub

Private Sub Opt_Funcao_Click(Index As Integer)
    MSFG_Provas.Rows = 1
    MSFG_Provas.Rows = 2
    MSFG_Provas.Enabled = False
    MebMatricula.PromptInclude = False
    MebMatricula.Text = ""
    MebMatricula.PromptInclude = True
    Lb_Aluno.Caption = ""
    Lb_ProvaE.Caption = ""
    Lb_Prova.Caption = ""
    Lb_Prova.Caption = ""
    Bt_ExclProva.Enabled = False
    Bt_GrvNotaS.Enabled = False
    Bt_GrvNotaS.Enabled = False
    
    Select Case Index
        Case 0
            Funcao = True
            Frame_ExclProva.Visible = False
            SistLanca
            Label3.Caption = "LANÇAR NOTAS"
            Label3.BackColor = vbBlue
            Frame_Sintetico.Visible = True
        Case 1
            Label3.Caption = "EXCLUIR PROVAS"
            Label3.BackColor = vbRed
            Funcao = False
            Frame_Sintetico.Visible = False
            'Frame_Analitico.Visible = False
            Frame_ExclProva.Visible = True
    End Select
End Sub
Private Sub LstProvas() 'Lista as provas que ja forao lançadas as nota/Excluir
    Dim bc As Integer 'cor de fundo da tela
    Dim EnsinoIDMatr As Integer
    Dim nmModulo        As String
    RefTrafegoID = Empty
    EnsinoID = Empty
    DisciplinaID = Empty
    nProva = Empty
    
    EnsinoIDMatr = PgMatrEnsino(MatrID, False) 'Pega o ensino nao concluido
    If EnsinoIDMatr = 0 Then
        MSFG_Provas.Rows = 1
        MsgBox "Matricula não possui ensino iniciado.", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    MSFG_Provas.Rows = 1
    'MSFG_Provas.Rows = 2
    Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoIDMatr & " AND Nota <> Null ORDER BY DisciplinaID ASC, NProva ASC")
    If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
            MsgBox "Nenhuma Prova para esta matricula!", vbInformation, "CESNet - Aviso!"
            MebMatricula.SetFocus
            MebMatricula_GotFocus
            Exit Sub
        Else
            RsMatriculaProva.MoveFirst
            lin = 1
            bc = 0
            Do Until RsMatriculaProva.EOF

                'Trocar a cor de fundo
                If DisciplinaID <> RsMatriculaProva.Fields("DisciplinaID") Then
                    If DisciplinaID <> 0 Then
                        bc = IIf(bc = 0, 8, 0)
                    End If
                End If
            
                RefTrafegoID = RsMatriculaProva.Fields("RefTrafegoID")
                EnsinoID = IIf(IsNull(RsMatriculaProva.Fields("EnsinoID")), 0, RsMatriculaProva.Fields("EnsinoID"))
                DisciplinaID = IIf(IsNull(RsMatriculaProva.Fields("DisciplinaID")), 0, RsMatriculaProva.Fields("DisciplinaID"))
                nProva = RsMatriculaProva.Fields("NProva")
                'Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "'")
                Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND NProva = '" & nProva & "'")
                If RsProva.BOF And RsProva.EOF Then
                        'MsgBox "Erro na Referencia da Prova chame o Suporte.", vbInformation, "CESNet - Aviso!"
                        'Exit Sub
                        'RefTrafegoID = 0
                        nmModulo = ""
                        Assunto = ""
                    Else
                        RsProva.MoveFirst
                        Assunto = IIf(IsNull(RsMatriculaProva.Fields("Assunto")), RsProva.Fields("Assunto"), RsMatriculaProva.Fields("Assunto"))
                        nmModulo = PgNomeModulo(RsProva.Fields("ModuloID"))
                End If
                MSFG_Provas.Rows = MSFG_Provas.Rows + 1
                MSFG_Provas.TextMatrix(lin, 0) = RefTrafegoID
                MSFG_Provas.TextMatrix(lin, 1) = PgNomeEnsino(RsMatriculaProva.Fields("EnsinoID"))
                MSFG_Provas.TextMatrix(lin, 2) = PgNomeDisciplina(DisciplinaID)
                MSFG_Provas.TextMatrix(lin, 3) = nmModulo
                MSFG_Provas.TextMatrix(lin, 4) = IIf(IsNull(RsMatriculaProva.Fields("DtAvaliacao")), " ", RsMatriculaProva.Fields("DtAvaliacao"))
                MSFG_Provas.TextMatrix(lin, 5) = nProva
                MSFG_Provas.TextMatrix(lin, 6) = Assunto
                MSFG_Provas.TextMatrix(lin, 7) = RsMatriculaProva.Fields("Tipo")
                MSFG_Provas.TextMatrix(lin, 8) = IIf(IsNull(RsMatriculaProva.Fields("Nota")), " ", RsMatriculaProva.Fields("Nota"))
                RsMatriculaProva.MoveNext
                'Troca a cor do fundo
                MSFG_Provas.Row = MSFG_Provas.Rows - 1
                MSFG_Provas.RowSel = MSFG_Provas.Rows - 1
                MSFG_Provas.Col = 0
                MSFG_Provas.ColSel = MSFG_Provas.Cols - 1
               
                MSFG_Provas.FillStyle = flexFillRepeat
                MSFG_Provas.CellBackColor = QBColor(bc)
                 
                           
                
                lin = lin + 1
            Loop
            
            'Alinha o Titulo das provas
            With MSFG_Provas
                
                If .Rows = 1 Then
                    .Rows = 2
                End If
                .Col = 5
                .ColSel = 5
                .Row = 1
                .RowSel = .Rows - 1
                .FillStyle = flexFillRepeat
                .CellAlignment = 1
                
                .Row = 1
            End With
    End If
    MSFG_Provas.Enabled = True
End Sub

Private Function PegSerie(RefTraf As Integer) As Integer
    Dim RsTrafego As Recordset
    Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE RefTrafegoID = " & RefTraf)
    If RsTrafego.BOF And RsTrafego.EOF Then
            PegSerie = 0
        Else
            RsTrafego.MoveFirst
            PegSerie = RsTrafego.Fields("SerieID")
    End If
    RsTrafego.Close
End Function

Private Sub Txt_ObsS_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
