VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Orientacao 
   Caption         =   "CESNet - Orientação"
   ClientHeight    =   6240
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   12705
   Icon            =   "Form_Orientacao.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6240
   ScaleWidth      =   12705
   Begin VB.Frame frmMenu 
      Height          =   4755
      Left            =   10200
      TabIndex        =   18
      Top             =   360
      Width           =   2295
      Begin VB.CheckBox chkHistProvas 
         Caption         =   "Mostrar historico de provas"
         Enabled         =   0   'False
         Height          =   435
         Left            =   180
         TabIndex        =   23
         Top             =   4020
         Width           =   1815
      End
      Begin VB.CommandButton Bt_HabAluno 
         Caption         =   "&Habilitar para Avaliação"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Form_Orientacao.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   22
         Top             =   180
         Width           =   2055
      End
      Begin VB.CommandButton Bt_OrientarAluno 
         Caption         =   "&Orientar Aluno"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Form_Orientacao.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1095
         Width           =   2055
      End
      Begin VB.CommandButton btFoto 
         Caption         =   "&Foto"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Form_Orientacao.frx":091E
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2940
         Width           =   2055
      End
      Begin VB.CommandButton btoAplicarProva 
         Caption         =   "&Aplicar Prova"
         Enabled         =   0   'False
         Height          =   855
         Left            =   120
         Picture         =   "Form_Orientacao.frx":0C28
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2025
         Width           =   2055
      End
   End
   Begin VB.Frame frmOrientacoes 
      Caption         =   "Ultimas Orientações:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2250
      Left            =   60
      TabIndex        =   14
      Top             =   3900
      Width           =   10035
      Begin VB.TextBox txtObs 
         Height          =   315
         Left            =   540
         MaxLength       =   254
         TabIndex        =   16
         Top             =   1800
         Width           =   9315
      End
      Begin MSFlexGridLib.MSFlexGrid MSFG_Orientacao 
         Height          =   1515
         Left            =   90
         TabIndex        =   15
         Top             =   225
         Width           =   9825
         _ExtentX        =   17330
         _ExtentY        =   2672
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_Orientacao.frx":0F32
      End
      Begin VB.Label lblObs 
         Alignment       =   1  'Right Justify
         Caption         =   "Obs.:"
         Height          =   195
         Left            =   120
         TabIndex        =   17
         Top             =   1860
         Width           =   375
      End
   End
   Begin VB.Frame frmProvas 
      Caption         =   "Provas:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1935
      Left            =   60
      TabIndex        =   12
      Top             =   1890
      Width           =   10035
      Begin MSFlexGridLib.MSFlexGrid MSFG_Provas 
         Height          =   1635
         Left            =   90
         TabIndex        =   13
         Top             =   225
         Width           =   9870
         _ExtentX        =   17410
         _ExtentY        =   2884
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_Orientacao.frx":1006
      End
   End
   Begin VB.Frame frmDados 
      Caption         =   "Dados do Aluno:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   60
      TabIndex        =   3
      Top             =   780
      Width           =   10035
      Begin VB.ComboBox Cb_Disciplina 
         Enabled         =   0   'False
         Height          =   315
         Left            =   7020
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   225
         Width           =   2700
      End
      Begin MSMask.MaskEdBox MebMatricula 
         Height          =   375
         Left            =   1035
         TabIndex        =   5
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
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Matricula:"
         Height          =   195
         Left            =   180
         TabIndex        =   11
         Top             =   315
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Disciplina:"
         Height          =   210
         Left            =   6240
         TabIndex        =   10
         Top             =   270
         Width           =   705
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   255
         Left            =   360
         TabIndex        =   9
         Top             =   720
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
         Left            =   1020
         TabIndex        =   8
         Top             =   660
         Width           =   8835
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Curso:"
         Height          =   195
         Left            =   3045
         TabIndex        =   7
         Top             =   315
         Width           =   555
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
         Left            =   3645
         TabIndex        =   6
         Top             =   225
         Width           =   2355
      End
   End
   Begin VB.Label lblTit 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "ORIENTAÇÃO"
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
      Width           =   12735
   End
   Begin VB.Label Lb_Professor 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   900
      TabIndex        =   1
      Top             =   420
      Width           =   8340
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Professor:"
      Height          =   195
      Left            =   60
      TabIndex        =   0
      Top             =   450
      Width           =   735
   End
End
Attribute VB_Name = "Form_Orientacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsProfessor             As Recordset
Dim RsMatriculaOrientacao   As Recordset
Dim RsProvas                As Recordset
Dim RsProvasTMP             As Recordset
'Dim RsProfessor             As Recordset

Dim MatrID                  As String
Dim EnsinoID                As Integer
Dim DisciplinaID            As Integer
Dim ProfID                  As Integer
Dim nProva                  As Integer
Dim Assunto                 As String
Dim DtAvaliacao             As String


Private Sub MstTodasProvas()
    Dim DtHrAv              As String
    Dim SerieID             As Integer
    Dim RefTrafegoID        As Integer
    Dim nmModulo            As String
    Dim nProva              As String
    Dim RsTrafego           As Recordset
    Dim RsProvas            As Recordset
    Dim RsMatriculaProvas   As Recordset
    Dim sSQL                As String
    Dim RsMatriculaSerie    As Recordset
    If EnsinoID = 0 Or DisciplinaID = 0 Then
        Exit Sub
    End If
    
    With MSFG_Provas
        .Rows = 1
    
    
        '##sSQL = "SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & _
                "' AND EnsinoID = " & EnsinoID & _
                " AND DisciplinaID = " & DisciplinaID & _
                IIf(chkHistProvas.Value = 0, " AND Aprovado = False ", " ") & _
                "ORDER BY SerieID"
        '##Set RsMatriculaSerie = BD.OpenRecordset(sSQL)
        '##If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
        '##        MsgBox "Não existe nenhuma prova para esta Disciplina.", vbInformation, "CESNet - Atenção"
        '##        Exit Sub
        '##    Else
        '##        RsMatriculaSerie.MoveFirst
        '##End If
        'Inicia o loop por Serie
        '##Do Until RsMatriculaSerie.EOF
            '##SerieID = RsMatriculaSerie.Fields("SerieID")
            'Pega o trafego
            
            '## = Removido 17.02.2012 Erro na listagen das provas
            '##Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)  'ModuloID = " & ModuloID)
            '##If RsTrafego.BOF And RsTrafego.EOF Then
            '##        MsgBox "Modulo nao encontrado no Trafego. Por favor verifique.", vbInformation, "CESNet - Aviso!"
            '##        Exit Sub
            '##    Else
            '##        RsTrafego.MoveFirst
            '##End If
        'Loop pelo TRAFEGO dentro de SERIE
            '##Do Until RsTrafego.EOF
                '##RefTrafegoID = RsTrafego.Fields("RefTrafegoID")
                'Pega as prova
                'sSQL = "SELECT * FROM Provas WHERE RefTrafegoID = " & RefTrafegoID & " ORDER BY NProva"
                sSQL = "SELECT * FROM Provas WHERE EnsinoID=" & EnsinoID & " AND DisciplinaID=" & DisciplinaID & " ORDER BY NProva"
                Set RsProvas = BD.OpenRecordset(sSQL)
                If RsProvas.BOF And RsProvas.EOF Then
                    MsgBox "Não existe nenhuma Prova cadastrada para essa Discip.rows-1a.", vbInformation, "CESNet - Aviso!"
                    Exit Sub
                End If
                RsProvas.MoveFirst
                'Loop Nas PROVAS
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
                
                    Set RsMatriculaProvas = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND NProva = '" & nProva & "'")
                    If RsMatriculaProvas.BOF And RsMatriculaProvas.EOF Then
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 0) = nProva
                            .TextMatrix(.Rows - 1, 1) = nmModulo
                            .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsProvas.Fields("Assunto")), "", RsProvas.Fields("Assunto"))
                            
                        Else
                        
                            RsMatriculaProvas.MoveFirst
                            '##If chkHistProvas.Value = 1 Or RsMatriculaProvas.Fields("Aprovado") = False Then
                            If chkHistProvas.Value = 1 Then
                            
                                'Checar se exits alguma prova ao qual foi reprovado
                                '##sSQL = "SELECT * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & nProva & "' ORDER BY Seq, NProva"
                                sSQL = "SELECT * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID=" & DisciplinaID & " AND NProva = '" & nProva & "' ORDER BY Seq, NProva"
                                Set RsProvasTMP = BD.OpenRecordset(sSQL)
                                If RsProvasTMP.BOF And RsProvasTMP.EOF Then
                                    Else
                                        RsProvasTMP.MoveFirst
                                            Do Until RsProvasTMP.EOF
                                            
                                                .Rows = .Rows + 1
                                                DtHrAv = IIf(IsNull(RsProvasTMP.Fields("DtAvaliacao")), " ", RsProvasTMP.Fields("DtAvaliacao"))
                                                DtHrAv = DtHrAv & IIf(IsNull(RsProvasTMP.Fields("DtHrAv")), "", Mid(RsProvasTMP.Fields("DtHrAv"), InStr(RsProvasTMP.Fields("DtHrAv"), " "), Len(RsProvasTMP.Fields("DtHrAv"))))
                                                .TextMatrix(.Rows - 1, 0) = RsProvasTMP.Fields("NProva")
                                                .TextMatrix(.Rows - 1, 1) = nmModulo
                                                .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsProvas.Fields("Assunto")), "", RsProvas.Fields("Assunto"))
                                                .TextMatrix(.Rows - 1, 3) = DtHrAv
                                                .TextMatrix(.Rows - 1, 4) = IIf(IsNull(RsProvasTMP.Fields("Tipo")), "", RsProvasTMP.Fields("Tipo"))
                                                .TextMatrix(.Rows - 1, 5) = PgNomeProf(RsProvasTMP.Fields("ProfIDN"))
                                                .TextMatrix(.Rows - 1, 6) = IIf(IsNull(RsProvasTMP.Fields("Obs")), "", RsProvasTMP.Fields("Obs"))
                                                .Row = .Rows - 1 '.rows-1
                                                .Col = 0
                                                .ColSel = .Cols - 1
                                                .FillStyle = flexFillRepeat
                                                .CellForeColor = vbRed '&HFF&
                                                '.Row = 0
                                                RsProvasTMP.MoveNext
                                            Loop
                                End If
                            End If
                            DtHrAv = IIf(IsNull(RsMatriculaProvas.Fields("DtAvaliacao")), " ", RsMatriculaProvas.Fields("DtAvaliacao"))
                            DtHrAv = DtHrAv & IIf(IsNull(RsMatriculaProvas.Fields("DtHrAv")), "", Mid(RsMatriculaProvas.Fields("DtHrAv"), InStr(RsMatriculaProvas.Fields("DtHrAv"), " "), Len(RsMatriculaProvas.Fields("DtHrAv"))))
                            .Rows = .Rows + 1
                            .TextMatrix(.Rows - 1, 0) = nProva
                            .TextMatrix(.Rows - 1, 1) = nmModulo
                            .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsProvas.Fields("Assunto")), "", RsProvas.Fields("Assunto"))
                            .TextMatrix(.Rows - 1, 3) = DtHrAv
                            .TextMatrix(.Rows - 1, 4) = IIf(IsNull(RsMatriculaProvas.Fields("Tipo")), " ", RsMatriculaProvas.Fields("Tipo"))
                            .TextMatrix(.Rows - 1, 5) = PgNomeProf(RsMatriculaProvas.Fields("ProfIDN"))
                            .TextMatrix(.Rows - 1, 6) = IIf(IsNull(RsMatriculaProvas.Fields("Obs")), "", RsMatriculaProvas.Fields("Obs"))
                    End If
                    RsProvas.MoveNext
                
                    '.Rows = .Rows + 1
                Loop 'LOOP DAS PROVAS
                '##RsTrafego.MoveNext
            '##Loop
            '##RsMatriculaSerie.MoveNext
        '##Loop 'LOOP DAS SERIE
        '.Rows = .Rows - 1
        'Alinha o Titulo das provas
        .Col = 2
        .ColSel = 2
        .Row = 1
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = 1
    
        'Organiza o grid pelo num. da prova
        '.Row = 1
        '.RowSel = .Rows - 1
        '.Col = 0
        '.ColSel = .Cols - 1
        '.FillStyle = flexFillRepeat
        '.Sort = 1
    End With
End Sub

Private Sub Bt_HabAluno_Click()
    If MatrID = "" Then Exit Sub
    If DisciplinaID = 0 Then Exit Sub
        Set RsMatriculaOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND HabDisciplina = False")
        If RsMatriculaOrientacao.BOF And RsMatriculaOrientacao.EOF Then
                MsgBox "Não existe Disciplina a ser liberada!", vbInformation, "CESNet - Aviso"
                RsMatriculaOrientacao.Close
                Exit Sub
                
            Else
                If MsgBox("Habilitar Matrícula n. " & MatrID & " para efetuar a prova ?", vbInformation + vbYesNo, "CESNet - Aviso!") = vbNo Then Exit Sub
                
                RsMatriculaOrientacao.MoveFirst
                DtAvaliacao = RsMatriculaOrientacao.Fields("DtAvaliacao")
                '################################################################################
                '### 27/01/2012
                'RsMatriculaOrientacao.Edit
                'RsMatriculaOrientacao.Fields("HabDisciplina") = True
                'RsMatriculaOrientacao.Fields("ProfOrientID") = ProfID
                'RsMatriculaOrientacao.Fields("DtOrientacao") = Date
                'RsMatriculaOrientacao.Update
                '################################################################################
                '==========================================================
                'RsMatriculaOrientacao.MoveLast
                
                'RsMatriculaOrientacao.MoveFirst
                RsMatriculaOrientacao.AddNew
                RsMatriculaOrientacao.Fields("MatrID") = MatrID
                RsMatriculaOrientacao.Fields("EnsinoID") = EnsinoID
                RsMatriculaOrientacao.Fields("DisciplinaID") = DisciplinaID
                RsMatriculaOrientacao.Fields("DtAvaliacao") = DtAvaliacao
                RsMatriculaOrientacao.Fields("HabDisciplina") = True
                RsMatriculaOrientacao.Fields("ProfOrientID") = ProfID
                RsMatriculaOrientacao.Fields("DtOrientacao") = Date
                RsMatriculaOrientacao.Fields("HabDisciplina") = True
                RsMatriculaOrientacao.Fields("Obs") = IIf(Trim(txtObs.Text) = "", Null, Trim(txtObs.Text))
                RsMatriculaOrientacao.Fields("DtHr") = Now
                RsMatriculaOrientacao.Update
                RsMatriculaOrientacao.Close
                '===========================================================
                
            
                MsgBox "Matrícula habilitada para nova avaliação.", vbInformation, "CESNet - Aviso"
                
               
        End If
    
   CarregarUltimasOrientacaoes
End Sub

Private Sub Bt_OrientarAluno_Click()
    If MatrID = "" Then Exit Sub
    If DisciplinaID = 0 Then Exit Sub
        Set RsMatriculaOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY Cont")
        If RsMatriculaOrientacao.BOF And RsMatriculaOrientacao.EOF Then
               DtAvaliacao = ""
            Else
                RsMatriculaOrientacao.MoveLast
                DtAvaliacao = IIf(IsNull(RsMatriculaOrientacao.Fields("DtAvaliacao")), " ", RsMatriculaOrientacao.Fields("DtAvaliacao"))
        End If
        RsMatriculaOrientacao.AddNew
        RsMatriculaOrientacao.Fields("MatrID") = MatrID
        RsMatriculaOrientacao.Fields("EnsinoID") = EnsinoID
        RsMatriculaOrientacao.Fields("DisciplinaID") = DisciplinaID
        RsMatriculaOrientacao.Fields("DtAvaliacao") = IIf(Trim(DtAvaliacao) = "", Null, DtAvaliacao)
        
        'Pegar a ultima informacao e reg abaixo
        RsMatriculaOrientacao.Fields("HabDisciplina") = pgUltOrientacao
        RsMatriculaOrientacao.Fields("ProfOrientID") = ProfID
        RsMatriculaOrientacao.Fields("DtOrientacao") = Date
        RsMatriculaOrientacao.Fields("Obs") = IIf(Trim(txtObs.Text) = "", Null, Trim(txtObs.Text))
        RsMatriculaOrientacao.Fields("DtHr") = Now
        RsMatriculaOrientacao.Update
        RsMatriculaOrientacao.Close
        MsgBox "Matrícula Orientada com sucesso.", vbInformation, "CESNet-Aviso"
    CarregarUltimasOrientacaoes
    'LpForm

End Sub
Private Function pgUltOrientacao() As Boolean
    Dim Rst As Recordset
    
    Set Rst = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY Cont")
    If Rst.BOF And Rst.EOF Then
            pgUltOrientacao = True
        Else
            Rst.MoveLast
            pgUltOrientacao = Rst.Fields("HabDisciplina")
    End If
    Rst.Close
End Function
Private Sub btFoto_Click()
    Form_ExibirImagem.ExibirFoto (MatrID)
End Sub

Private Sub btoAplicarProva_Click()
    If Trim(Cb_Disciplina.Text) = "" Then
        MsgBox "Selecione uma disciplina!", vbInformation, "Aviso"
        Exit Sub
    End If
    Form_Avaliacao.ReceberInformacoes MatrID, Cb_Disciplina.Text
End Sub

Private Sub Cb_Disciplina_Click()
    DisciplinaID = PgIDDisciplina(Trim(Cb_Disciplina.Text))
    MstTodasProvas
    CarregarUltimasOrientacaoes
    Exit Sub

End Sub


Private Function PgAssuntoProva(TrafID As Integer, nProva As String) As String
    Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & TrafID & " AND NProva = '" & nProva & "'")
    If RsProvas.BOF And RsProvas.EOF Then
            PgAssuntoProva = "<Assunto não localizado>"
        Else
            RsProvas.MoveFirst
            PgAssuntoProva = RsProvas.Fields("Assunto")
    End If
    RsProvas.Close
End Function

Private Function PgNomeProf(ProfID As Integer) As String
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores WHERE ProfID = " & ProfID)
    If RsProfessor.BOF And RsProfessor.EOF Then
            PgNomeProf = "  "
        Else
            RsProfessor.MoveFirst
            PgNomeProf = RsProfessor.Fields("Nome")
    End If
    RsProfessor.Close
End Function

Private Sub Cb_Disciplina_DropDown()
    Cb_Disciplina.Clear
    'Set RsMatriculaOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND HabDisciplina = False ORDER BY DtAvaliacao")
    Set RsMatriculaOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND IsNull(DtConclusao)")
    If RsMatriculaOrientacao.BOF And RsMatriculaOrientacao.EOF Then
            Cb_Disciplina.Clear
            Bt_HabAluno.Enabled = False
            Bt_OrientarAluno.Enabled = False
            btoAplicarProva.Enabled = False
            btFoto.Enabled = False
            Exit Sub
        Else
            RsMatriculaOrientacao.MoveFirst
            Do Until RsMatriculaOrientacao.EOF
                Cb_Disciplina.AddItem (PgNomeDisciplina(RsMatriculaOrientacao.Fields("DisciplinaID")))
                RsMatriculaOrientacao.MoveNext
            Loop
            Bt_HabAluno.Enabled = True
            Bt_OrientarAluno.Enabled = True
            btoAplicarProva.Enabled = True
            btFoto.Enabled = True
    End If
    RsMatriculaOrientacao.Close
End Sub


Private Sub chkHistProvas_Click()
    MstTodasProvas
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores WHERE UsuarioID = " & UsuarioID)
    If RsProfessor.BOF And RsProfessor.EOF Then
            MsgBox "Usuário não cadastrado para esta sessão.", vbInformation, "CESNet - Aviso!"
            Exit Sub
            Unload Me
            
        Else
            RsProfessor.MoveFirst
            ProfID = RsProfessor.Fields("ProfID")
            Lb_Professor.Caption = RsProfessor.Fields("Nome")
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    lblTit.Width = Me.Width
    
    frmMenu.left = Me.Width - (frmMenu.Width + 150)
    
    frmProvas.Width = Me.Width - (frmMenu.Width + 150 + 150)
    frmProvas.Height = (Me.Height - (frmDados.Height + 1600)) / 2
    MSFG_Provas.Width = frmProvas.Width - 200
    MSFG_Provas.Height = frmProvas.Height - 300
    
    frmOrientacoes.top = frmProvas.top + frmProvas.Height + 150
    frmOrientacoes.Width = frmProvas.Width
    frmOrientacoes.Height = frmProvas.Height
    MSFG_Orientacao.Width = frmOrientacoes.Width - 200
    MSFG_Orientacao.Height = frmOrientacoes.Height - (300 + txtObs.Height)
    
    txtObs.top = MSFG_Orientacao.top + MSFG_Orientacao.Height
    txtObs.Width = MSFG_Orientacao.Width - (lblObs.Width + 100)
    lblObs.top = txtObs.top
    
    
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
    End If
End Sub
Private Sub MstDadosAluno()


    Lb_Nome.Caption = PgDadosMatr(MatrID).Nome
        EnsinoID = PgMatrEnsino(MatrID, False)
        If EnsinoID = 0 Then
                btFoto.Enabled = False
                Lb_Ensino.Caption = ""
                Lb_Nome.Caption = ""
                Cb_Disciplina.Clear
                Cb_Disciplina.Enabled = False
                chkHistProvas.Enabled = False
            Else
                btFoto.Enabled = True
                Lb_Ensino.Caption = PgNomeEnsino(EnsinoID)
                Cb_Disciplina.Enabled = True
                Cb_Disciplina.Clear
                chkHistProvas.Enabled = True
                MSFG_Provas.Rows = 1
                MSFG_Orientacao.Rows = 1
        End If
        Bt_HabAluno.Enabled = False
        Bt_OrientarAluno.Enabled = False
        btoAplicarProva.Enabled = False
        
End Sub
Private Sub LpForm()
    MatrID = ""
    MebMatricula.PromptInclude = False
    MebMatricula.Text = ""
    MebMatricula.PromptInclude = True
    
    Lb_Ensino.Caption = ""
    
    Cb_Disciplina.Clear
    
    MSFG_Provas.Rows = 1
    MSFG_Provas.Rows = 2
    MSFG_Orientacao.Rows = 1
    MSFG_Orientacao.Rows = 2
    
    Lb_Nome.Caption = ""
End Sub
Private Sub CarregarUltimasOrientacaoes()
    Dim RsUO As Recordset
    MSFG_Orientacao.Rows = 1
    MSFG_Orientacao.Rows = 2
    Set RsUO = BD.OpenRecordset("SELECT TOP 100 * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY Cont DESC")
    If RsUO.BOF And RsUO.EOF Then
            RsUO.Close
            Exit Sub
        Else
            MSFG_Orientacao.Rows = 1
            RsUO.MoveFirst
            Do Until RsUO.EOF
                MSFG_Orientacao.Rows = MSFG_Orientacao.Rows + 1
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 0) = RsUO.Fields("Cont")
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 1) = PgNomeDisciplina(DisciplinaID)
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 2) = IIf(IsNull(RsUO.Fields("DtAvaliacao")), "", RsUO.Fields("DtAvaliacao"))
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 3) = IIf(IsNull(RsUO.Fields("DtOrientacao")), "", RsUO.Fields("DtOrientacao"))
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 4) = PgNomeProfessor(IIf(IsNull(RsUO.Fields("ProfOrientID")), 0, RsUO.Fields("ProfOrientID")))
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 5) = IIf(RsUO.Fields("HabDisciplina") = True, "SIM", "NÃO")
                MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Rows - 1, 6) = IIf(IsNull(RsUO.Fields("Obs")) = True, "", RsUO.Fields("Obs"))
                RsUO.MoveNext
            Loop
    End If
End Sub


Private Sub MSFG_Orientacao_Click()
    If MSFG_Orientacao.MouseRow = 0 Then Exit Sub
    txtObs.Text = MSFG_Orientacao.TextMatrix(MSFG_Orientacao.Row, 6)
End Sub
