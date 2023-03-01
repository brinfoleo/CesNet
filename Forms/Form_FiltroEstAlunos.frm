VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_FiltroEstAlunos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Filtro de Estatistica de Alunos"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8130
   Icon            =   "Form_FiltroEstAlunos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   8130
   Begin VB.Frame Frame1 
      Height          =   4455
      Left            =   60
      TabIndex        =   3
      Top             =   360
      Width           =   5835
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Cursando na Disciplina"
         Height          =   195
         Index           =   7
         Left            =   2760
         TabIndex        =   21
         Top             =   600
         Width           =   2235
      End
      Begin VB.Frame frmPeriodo 
         Caption         =   "Periodo:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   195
         TabIndex        =   15
         Top             =   2175
         Width           =   2865
         Begin MSComCtl2.DTPicker DTP_Inicial 
            Height          =   315
            Left            =   1080
            TabIndex        =   16
            Top             =   300
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   16711680
            CalendarTitleForeColor=   16777215
            Format          =   56295425
            CurrentDate     =   38393
         End
         Begin MSComCtl2.DTPicker DTP_Final 
            Height          =   315
            Left            =   1080
            TabIndex        =   17
            Top             =   780
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   556
            _Version        =   393216
            CalendarTitleBackColor=   16711680
            CalendarTitleForeColor=   16777215
            Format          =   56295425
            CurrentDate     =   38393
         End
         Begin VB.Label Label2 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Final:"
            Height          =   195
            Left            =   180
            TabIndex        =   19
            Top             =   840
            Width           =   795
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            Caption         =   "Data Inicial:"
            Height          =   195
            Left            =   120
            TabIndex        =   18
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Frequencia"
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   14
         Top             =   300
         Value           =   -1  'True
         Width           =   2235
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Concluido por Disciplina"
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   13
         Top             =   820
         Width           =   2235
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Concluidos por Ensino"
         Height          =   195
         Index           =   2
         Left            =   180
         TabIndex        =   12
         Top             =   1080
         Width           =   2235
      End
      Begin VB.Frame Frame_Criterio 
         Caption         =   "Criterio:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   780
         Left            =   3240
         TabIndex        =   9
         Top             =   2235
         Visible         =   0   'False
         Width           =   1815
         Begin VB.OptionButton Opt_Criterio 
            Caption         =   "Analitico"
            Height          =   195
            Index           =   1
            Left            =   180
            TabIndex        =   11
            Top             =   495
            Width           =   1410
         End
         Begin VB.OptionButton Opt_Criterio 
            Caption         =   "Sintético"
            Height          =   195
            Index           =   0
            Left            =   180
            TabIndex        =   10
            Top             =   180
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Concluintes por Ensino"
         Height          =   195
         Index           =   3
         Left            =   180
         TabIndex        =   8
         Top             =   1340
         Width           =   2235
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Ativos e Inativos"
         Height          =   195
         Index           =   4
         Left            =   180
         TabIndex        =   7
         Top             =   1600
         Width           =   2235
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Iniciados na Disciplina"
         Height          =   195
         Index           =   5
         Left            =   180
         TabIndex        =   6
         Top             =   560
         Width           =   2235
      End
      Begin VB.OptionButton Opt_Estatistica 
         Caption         =   "Matricula com renovação vencida"
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   4
         Top             =   1860
         Width           =   2895
      End
      Begin ComctlLib.ProgressBar pb 
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   4155
         Width           =   5475
         _ExtentX        =   9657
         _ExtentY        =   344
         _Version        =   327682
         Appearance      =   1
      End
      Begin VB.Label lblTipoRelatorio 
         Caption         =   "Label3"
         Height          =   435
         Left            =   240
         TabIndex        =   20
         Top             =   3555
         Width           =   5415
      End
   End
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   750
      Left            =   6000
      Picture         =   "Form_FiltroEstAlunos.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1200
      Width           =   2055
   End
   Begin VB.CommandButton Bt_Aplicar 
      Caption         =   "&Aplicar"
      Height          =   765
      Left            =   6000
      Picture         =   "Form_FiltroEstAlunos.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   420
      Width           =   2055
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "FILTRO DE ESTATISTICA DE ALUNOS"
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
      Width           =   8175
   End
End
Attribute VB_Name = "Form_FiltroEstAlunos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim RsMatr As Recordset
Dim RsMatrEnsino As Recordset
Dim RsMatrDisciplina As Recordset
Dim RsMatrProva As Recordset
Dim Crit As Integer
Dim Opcao As Integer ' E a opcao de relatorio selecionado


Private Sub GerarArquivo(txt As String, LocalArquivo As String)
'Apagar apos uso Esquema Miguel

    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    'Dim msg As String
    'Dim caminho As String
    'caminho = App.Path & "\ErrLog.txt"
'se o arquivo não existir então cria
    If fso.FileExists(LocalArquivo) Then
            Set Arquivo = fso.GetFile(LocalArquivo)
        Else
            Set arquivoLog = fso.CreateTextFile(LocalArquivo)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(LocalArquivo)
    End If
'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
'monta informações para gerar a linha com erro
    'msg = "[" & Now & "]" & Form & ":[" & Num & "-" & Descr & "]"
' inclui linhas no arquivo texto
    arquivoLog.WriteLine txt
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)
'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing

End Sub

Private Sub Bt_Aplicar_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    Select Case Opcao
        Case 0
            Call RptFreqAluno
        Case 1 'Concluido por Disciplina
            Call RptConclDisc
        Case 2
            Call RptConclEnsino
        Case 3
            Call RptConclIntEnsino
        Case 4
            Call RptAtivoInativo
        Case 5 'Iniciado por Disciplina
            Call RptInicDisc
        Case 6
            Call RptRenovacaoVencida
        Case 7 'Cursando na Disciplina
            Call RptCursandoDisc
    End Select
End Sub
Private Sub RptRenovacaoVencida()
    'Lista as MAtriculas com data de renovação vencida
    Dim sSQL As String
    Dim Rst As Recordset
    
    
    sSQL = "SELECT MatriculaEnsino.MatrID, Matriculas.Nome, MatriculaEnsino.DtInicio, MatriculaEnsino.DtRenovacao, MatriculaEnsino.Trancado, MatriculaEnsino.DtFinal " & _
           "FROM MatriculaEnsino INNER JOIN Matriculas ON MatriculaEnsino.MatrID = Matriculas.MatrID " & _
           "WHERE (((MatriculaEnsino.DtRenovacao)<=#" & Format(Date, "MM/DD/YYYY") & "#) AND ((MatriculaEnsino.Trancado)=False) AND ((MatriculaEnsino.DtFinal) Is Null))"
           
           
           
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Nenhuma Renovação vencida!", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            'rptListRenovacaoVencida.DataSource = Rst.Data
            Call Relatorio(rptListRenovacaoVencida, sSQL)
            'rptListAtivosInativos.Sections("Section2").Controls("lblAtivos").Caption = cAtivo
            'rptListAtivosInativos.Sections("Section2").Controls("lblInativos").Caption = cInativo
            rptListRenovacaoVencida.Show 1
    End If
    Rst.Close
End Sub
Private Sub ImprDados(Dt As Date, Provas As Integer)
    'If Trim(dtAv) = "0" Then Exit Sub
    DoEvents
    ObjPreview.FontSize = CI.tFonte
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Sublinhado
                
    ObjPreview.Print Tab(5); Dt; _
                     Tab(30); left("000", 3 - Len(Trim(Provas))) & Trim(Provas)
End Sub
Private Sub Cab1()
    DoEvents
    Call cPreview(2)
    ObjPreview.FontBold = True
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO DE FREQUENCIA DE ALUNOS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value) / 2)
    ObjPreview.Print "RELATÓRIO DE FREQUENCIA DE ALUNOS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    ObjPreview.Print
    ObjPreview.Print Tab(5); "Data"; _
                     Tab(25); "Num. de Alunos"; ' _
                     Tab(40); "Nome do Aluno"; _
                     Tab(100); "Qtd. provas na data"
End Sub
Private Sub Cab2()
    DoEvents
    Call cPreview(2)
    ObjPreview.FontBold = True
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO DE DISCIPLINAS CONCLUIDAS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value) / 2)
    ObjPreview.Print "RELATÓRIO DE DISCIPLINAS CONCLUIDAS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    ObjPreview.Print
    'ObjPreview.Print Tab(5); "Data"; _
                     Tab(25); "Num. de Provas"; ' _
                     Tab(40); "Nome do Aluno"; _
                     Tab(100); "Qtd. provas na data"
End Sub
Private Sub Cab3()
    DoEvents
    Call cPreview(2)
    ObjPreview.FontBold = True
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO DE ENSINOS CONCLUIDOS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value) / 2)
    ObjPreview.Print "RELATÓRIO DE ENSINOS CONCLUIDOS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    ObjPreview.Print
    'ObjPreview.Print Tab(5); "Data"; _
                     Tab(25); "Num. de Provas"; ' _
                     Tab(40); "Nome do Aluno"; _
                     Tab(100); "Qtd. provas na data"
End Sub
Private Sub Cab4()
    DoEvents
    Call cPreview(2)
    ObjPreview.FontBold = True
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO DE ENSINOS CONCLUINTES NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value) / 2)
    ObjPreview.Print "RELATÓRIO DE ENSINOS CONCLUINTES NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    ObjPreview.Print
    'ObjPreview.Print Tab(5); "Data"; _
                     Tab(25); "Num. de Provas"; ' _
                     Tab(40); "Nome do Aluno"; _
                     Tab(100); "Qtd. provas na data"
End Sub
Private Sub Cab5()
    DoEvents
    Call cPreview(2)
    ObjPreview.FontBold = True
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO DOS ALUNOS ATIVOS E INATIVOS NOS ULTIMOS " & InatDias & " DIAS") / 2)
    ObjPreview.Print "RELATÓRIO DOS ALUNOS ATIVOS E INATIVOS NOS ULTIMOS " & InatDias & " DIAS"
    ObjPreview.Print
    'ObjPreview.Print Tab(5); "Data"; _
                     Tab(25); "Num. de Provas"; ' _
                     Tab(40); "Nome do Aluno"; _
                     Tab(100); "Qtd. provas na data"
End Sub
Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If

End Sub

'Private Function PgTotProvas(m As String, dt As String) As String
    'Erro na quantidade de provas
'    Dim RsTMP As Recordset
'    Dim Prvs As String
'    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & m & "' AND DtAvaliacao = #" & Format(dt, "MM/DD/YYYY") & "#")
'    If RsTMP.BOF And RsTMP.EOF Then
'            MsgBox "Erro ao localizar Matricula e Qtd de provas", vbInformation, "CESNet - Aviso"
'            PgTotProvas = "000"
'        Else
'            RsTMP.MoveLast
'            PgTotProvas = RsTMP.RecordCount
'    End If
'End Function
Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    DTP_Final.Value = Date
    DTP_Inicial.Value = DTP_Final - 30
    lblTipoRelatorio.Caption = ""
End Sub

Private Sub Opt_Criterio_Click(Index As Integer)
    Crit = Index
End Sub

Private Sub Opt_Estatistica_Click(Index As Integer)
    Frame_Criterio.Visible = False
    'DTP_Inicial.Enabled = True
    'DTP_Final.Enabled = True
    frmPeriodo.Visible = True
    Select Case Index
        Case 0
            Opcao = 0
        Case 1
            Opcao = 1
        Case 2
            Frame_Criterio.Visible = True
            Crit = 0
            Opcao = 2
        Case 3
            Frame_Criterio.Visible = True
            Crit = 0
            Opcao = 3
        Case 4
            'Frame_Criterio.Visible = True
            Crit = 0
            'DTP_Inicial.Enabled = False
            'DTP_Final.Enabled = False
            frmPeriodo.Visible = False
            Opcao = 4
            lblTipoRelatorio.Caption = "Matriculas com curso iniciado..."
        Case 5
            Opcao = 5
        Case 6
            frmPeriodo.Visible = False
            Opcao = 6
            lblTipoRelatorio.Caption = "Matriculas com prazo de renovação vencidas..."
        Case 7 ' Cursando na Disciplina
            Opcao = 7
            
    End Select
End Sub
Private Function ChecarData() As Boolean
    If DTP_Final.Value < DTP_Inicial.Value Then
            ChecarData = False
            MsgBox "A Data Final não deve ser inferior a Data Inicial.", vbInformation, "CESNet - Aviso!"
            Exit Function
        Else
            ChecarData = True
    End If
End Function

Private Sub RptInicDisc()
    Dim Criterio As String
    
    Criterio = "SELECT Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio, Count(MatriculaDisciplina.MatrID) AS ContarDeMatrID " & _
                "FROM (MatriculaDisciplina INNER JOIN Ensino ON MatriculaDisciplina.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaDisciplina.DisciplinaID = Disciplina.ID " & _
                "GROUP BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio " & _
                "Having (((MatriculaDisciplina.DtInicio) >= #" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# And (MatriculaDisciplina.DtInicio) <= #" & Format(DTP_Final.Value, "MM/DD/YYYY") & "#)) " & _
                "ORDER BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio"



    Call Relatorio(rptDisciplIniciadas, Criterio)
    rptDisciplIniciadas.Sections("Cab").Controls("lbPeriodo").Caption = "PERÍODO DE " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    rptDisciplIniciadas.Show
    
End Sub
Private Sub RptCursandoDisc()
    Dim Criterio As String
    
'    Criterio = "SELECT Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio, MatriculaDisciplina.DtConclusao, Count(MatriculaDisciplina.MatrID) AS ContarDeMatrID " & _
                "FROM (MatriculaDisciplina INNER JOIN Ensino ON MatriculaDisciplina.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaDisciplina.DisciplinaID = Disciplina.ID " & _
                "GROUP BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio " & _
                "Having (((MatriculaDisciplina.DtInicio) >= #" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# And (MatriculaDisciplina.DtInicio) <= #" & Format(DTP_Final.Value, "MM/DD/YYYY") & "#) AND MatriculaDisciplina.DtConclusao = ISNULL) " & _
                "ORDER BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio"

    Criterio = "SELECT Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio, MatriculaDisciplina.DtConclusao, Count(MatriculaDisciplina.MatrID) AS ContarDeMatrID " & _
                "FROM (MatriculaDisciplina INNER JOIN Ensino ON MatriculaDisciplina.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaDisciplina.DisciplinaID = Disciplina.ID " & _
                "GROUP BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtInicio, MatriculaDisciplina.DtConclusao " & _
                "HAVING (((MatriculaDisciplina.DtInicio)>=#" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# And (MatriculaDisciplina.DtInicio)<=#" & Format(DTP_Final.Value, "MM/DD/YYYY") & "#) AND ((MatriculaDisciplina.DtConclusao) Is Null))"



    Call Relatorio(rptDisciplIniciadas, Criterio)
    rptDisciplIniciadas.Sections("Cab").Controls("lbPeriodo").Caption = "PERÍODO DE " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    rptDisciplIniciadas.Show
    
End Sub

Private Sub RptConclDisc()
    Dim Criterio As String
    
    'Criterio = "SELECT Ensino.Descr, Disciplina.Descr, Count(MatriculaDisciplina.MatrID) AS ContarDeMatrID " & _
    '         "FROM (MatriculaDisciplina INNER JOIN Ensino ON MatriculaDisciplina.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaDisciplina.DisciplinaID = Disciplina.ID " & _
    '         "GROUP BY Ensino.Descr, Disciplina.Descr " & _
    '         "ORDER BY Ensino.Descr, Disciplina.Descr"
    Criterio = "SELECT Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtConclusao, Count(MatriculaDisciplina.MatrID) AS ContarDeMatrID " & _
                "FROM (MatriculaDisciplina INNER JOIN Ensino ON MatriculaDisciplina.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaDisciplina.DisciplinaID = Disciplina.ID " & _
                "GROUP BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtConclusao " & _
                "Having (((MatriculaDisciplina.DtConclusao) >= #" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# And (MatriculaDisciplina.DtConclusao) <= #" & Format(DTP_Final.Value, "MM/DD/YYYY") & "#)) " & _
                "ORDER BY Ensino.Descr, Disciplina.Descr, MatriculaDisciplina.DtConclusao"


    Call Relatorio(rptDisciplConcluidas, Criterio)
    rptDisciplConcluidas.Sections("Cab").Controls("lbPeriodo").Caption = "PERÍODO DE " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    rptDisciplConcluidas.Show
    
    
    'Dim AlCur(30, 30) As Integer
    'Dim tmp As Integer
    'Dim cont As Integer
    'Dim tot As Integer
    'If ChecarData = False Then Exit Sub
    'Set RsMatrDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE DtConclusao >= #" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# AND DtConclusao <= #" & Format(DTP_Final.Value, "MM/DD/YYYY") & "# ORDER BY EnsinoID")
    'If RsMatrDisciplina.BOF And RsMatrDisciplina.EOF Then
    '        MsgBox "Nenhum Aluno Cursando neste Periodo", vbInformation, "CESNet - Aviso"
    '        Exit Sub
    '    Else
    '        RsMatrDisciplina.MoveFirst
    'End If
    'Do Until RsMatrDisciplina.EOF
    '    AlCur(RsMatrDisciplina.Fields("EnsinoID"), RsMatrDisciplina.Fields("DisciplinaID")) = AlCur(RsMatrDisciplina.Fields("EnsinoID"), RsMatrDisciplina.Fields("DisciplinaID")) + 1
    '    RsMatrDisciplina.MoveNext
    'Loop
    'If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
    '    Exit Sub
    'End If
    'Call Cab2
    'DoEvents
    'ObjPreview.FontSize = CI.tFonte
    'ObjPreview.Font = CI.Fonte
    'ObjPreview.FontBold = CI.Negrito
    'ObjPreview.FontItalic = CI.Italico
    'ObjPreview.FontUnderline = CI.Sublinhado
    'For tmp = 1 To 30
    '    If PgNomeEnsino(tmp) = 0 Then
    '        Else
    '            For cont = 1 To 30
    '                If PgNomeDisciplina(cont) = 0 Then
    '                    Else
    '                        ObjPreview.Print Tab(5); PgNomeEnsino(tmp) & " / "; PgNomeDisciplina(cont); Tab(70); " = "; left("000", 3 - Len(Trim(AlCur(tmp, cont)))) & Trim(AlCur(tmp, cont))
    '                        tot = tot + AlCur(tmp, cont)
    '                End If
    '            Next
    '            ObjPreview.FontBold = True
    '            ObjPreview.Print Tab(5); "TOTAL: "; Tab(15); tot
    '            ObjPreview.FontBold = CI.Negrito
    '            ObjPreview.Print
    '            tot = 0
    '    End If
    'Next
End Sub
Private Sub RptFreqAluno()
    Dim QtdProvas As Integer
    Dim QtdProvasDia As Integer
    
    Dim DtAnt As Date
    'QtdProvas = 0
    'QtdProvasDia = 0
    
    If ChecarData = False Then Exit Sub
    Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE DtAvaliacao >= #" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# AND DtAvaliacao <= #" & Format(DTP_Final.Value, "MM/DD/YYYY") & "# ORDER BY DtAvaliacao,MatrID")
    If RsMatrProva.BOF And RsMatrProva.EOF Then
            MsgBox "Nenhuma prova encontrada.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            
            RsMatrProva.MoveFirst
            If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
                Exit Sub
            End If
            Call Cab1
            Dim tmpMatr As String
            QtdProvasDia = 1
            tmpMatr = RsMatrProva.Fields("MatrID")
            DtAnt = RsMatrProva.Fields("DtAvaliacao")
            Do Until RsMatrProva.EOF
                If DtAnt = RsMatrProva.Fields("DtAvaliacao") Then
                        If tmpMatr = RsMatrProva.Fields("MatrID") Then
                            Else
                                tmpMatr = RsMatrProva.Fields("MatrID")
                                QtdProvasDia = QtdProvasDia + 1
                        End If
                        
                        RsMatrProva.MoveNext
                        
                        If RsMatrProva.EOF Then
                            Call ImprDados(DtAnt, QtdProvasDia)
                            QtdProvas = QtdProvas + QtdProvasDia
                        End If
                    Else
                        Call ImprDados(DtAnt, QtdProvasDia)
                        QtdProvas = QtdProvas + QtdProvasDia
                        QtdProvasDia = 1
                        DtAnt = RsMatrProva.Fields("DtAvaliacao")
                        RsMatrProva.MoveNext
                End If
            Loop
            ObjPreview.Print
            ObjPreview.FontBold = True
            ObjPreview.Print Tab(5); "TOTAL: "; QtdProvas
    End If
End Sub
Private Sub RptConclIntEnsino()
    
    If ChecarData = False Then Exit Sub
    Dim ConcEnsi(2, 30) As Integer
    Dim tmp As Integer
    Dim Sexo As Integer
    Set RsMatrEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE DtInicio >= #" & _
                       Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# AND DtInicio <= #" & _
                       Format(DTP_Final.Value, "MM/DD/YYYY") & "# AND IsNull(DtFinal)" & _
                       " ORDER BY EnsinoID,DtFinal")
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
            MsgBox "Nenhuma Disciplina Concluida no Periodo.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsMatrEnsino.MoveFirst
    End If
   
    If Crit = 0 Then
            'Sintetico
            Do Until RsMatrEnsino.EOF
                Sexo = IIf(PgDadosMatr(RsMatrEnsino.Fields("MatrID")).Sexo = "M", 1, 2)
                ConcEnsi(Sexo, RsMatrEnsino.Fields("EnsinoID")) = ConcEnsi(Sexo, RsMatrEnsino.Fields("EnsinoID")) + 1
                RsMatrEnsino.MoveNext
            Loop
            If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
                Exit Sub
            End If
            Call Cab4
            DoEvents
            ObjPreview.FontSize = CI.tFonte
            ObjPreview.Font = CI.Fonte
            ObjPreview.FontBold = CI.Negrito
            ObjPreview.FontItalic = CI.Italico
            ObjPreview.FontUnderline = CI.Sublinhado
            ObjPreview.Print Tab(10); "ENSINO:"
            ObjPreview.Print
            For tmp = 1 To 30
                If PgNomeEnsino(tmp) = 0 Then
                    Else
                        ObjPreview.Print Tab(15); PgNomeEnsino(tmp); " (Masc.) = "; ConcEnsi(1, tmp)
                        ObjPreview.Print Tab(15); PgNomeEnsino(tmp); " (Fem.) = "; ConcEnsi(2, tmp)
                End If
            Next
        Else
            'Analitico
            If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
                Exit Sub
            End If
            Call Cab4
            DoEvents
            ObjPreview.FontSize = CI.tFonte
            ObjPreview.Font = CI.Fonte
            ObjPreview.FontBold = CI.Negrito
            ObjPreview.FontItalic = CI.Italico
            ObjPreview.FontUnderline = CI.Sublinhado
            Do Until RsMatrEnsino.EOF
                ObjPreview.Print Tab(5); RsMatrEnsino.Fields("MatrID"); _
                                 Tab(20); PgDadosMatr(RsMatrEnsino.Fields("MatrID")).Nome; _
                                 Tab(80); PgNomeEnsino(RsMatrEnsino.Fields("EnsinoID")); _
                                 Tab(100); RsMatrEnsino.Fields("DtInicio");
                                 'Tab(115); RsMatrEnsino.Fields("Local")
                RsMatrEnsino.MoveNext
            Loop
    End If
End Sub



Private Sub RptConclEnsino()
    
    If ChecarData = False Then Exit Sub
    
    Dim ConcEnsi(2, 30) As Integer
    Dim tmp             As Integer
    Dim Sexo            As Integer '1 - masc // 2 - fem
    Dim sSQL            As String
    
    sSQL = "SELECT * FROM MatriculaEnsino WHERE DtFinal >= #" & _
                        Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# AND DtFinal <= #" & _
                        Format(DTP_Final.Value, "MM/DD/YYYY") & "# " & _
                        "AND Trancado = FALSE " & _
                        "ORDER BY EnsinoID,DtFinal"
    
    Set RsMatrEnsino = BD.OpenRecordset(sSQL)
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
            MsgBox "Nenhuma Disciplina Concluida no Periodo", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsMatrEnsino.MoveFirst
    End If
   
    If Crit = 0 Then
            'Sintetico
            Do Until RsMatrEnsino.EOF
                Sexo = IIf(PgDadosMatr(RsMatrEnsino.Fields("MatrID")).Sexo = "M", 1, 2)
                ConcEnsi(Sexo, RsMatrEnsino.Fields("EnsinoID")) = ConcEnsi(Sexo, RsMatrEnsino.Fields("EnsinoID")) + 1
                RsMatrEnsino.MoveNext
            Loop
            If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
                Exit Sub
            End If
            Call Cab3
            DoEvents
            ObjPreview.FontSize = CI.tFonte
            ObjPreview.Font = CI.Fonte
            ObjPreview.FontBold = CI.Negrito
            ObjPreview.FontItalic = CI.Italico
            ObjPreview.FontUnderline = CI.Sublinhado
            ObjPreview.Print Tab(10); "ENSINO:"
            ObjPreview.Print
            For tmp = 1 To 30
                If PgNomeEnsino(tmp) = 0 Then
                    Else
                        ObjPreview.Print Tab(15); PgNomeEnsino(tmp); " (Masc.) = "; ConcEnsi(1, tmp)
                        ObjPreview.Print Tab(15); PgNomeEnsino(tmp); " (Fem.) = "; ConcEnsi(2, tmp)
                End If
            Next
        Else
            'Analitico
            If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
                Exit Sub
            End If
            Call Cab3
            DoEvents
            ObjPreview.FontSize = CI.tFonte
            ObjPreview.Font = CI.Fonte
            ObjPreview.FontBold = CI.Negrito
            ObjPreview.FontItalic = CI.Italico
            ObjPreview.FontUnderline = CI.Sublinhado
            Do Until RsMatrEnsino.EOF
                ObjPreview.Print Tab(5); RsMatrEnsino.Fields("MatrID"); _
                                 Tab(20); PgDadosMatr(RsMatrEnsino.Fields("MatrID")).Nome; _
                                 Tab(80); PgNomeEnsino(RsMatrEnsino.Fields("EnsinoID")); _
                                 Tab(100); RsMatrEnsino.Fields("DtFinal"); _
                                 Tab(115); RsMatrEnsino.Fields("Local")
                RsMatrEnsino.MoveNext
            Loop
    End If
End Sub

Private Sub RptAtivoInativo()

    Dim sSQL        As String
    Dim cAtivo      As Long
    Dim cInativo    As Long
    Dim Status      As String
    
    sSQL = "SELECT * FROM MatriculaEnsino WHERE" & _
                        " DtInicio <> Null AND IsNull(DtFinal)" & _
                        " AND Trancado=False " & _
                        " ORDER BY StatusMatr ASC"
    
    Set RsMatrEnsino = BD.OpenRecordset(sSQL)
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
            MsgBox "Nenhum Ensino em aberto.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            cAtivo = 0
            cInativo = 0
            pb.Value = 0
            pb.Min = 0
            pb.Max = RsMatrEnsino.RecordCount
    
            RsMatrEnsino.MoveFirst
            Do Until RsMatrEnsino.EOF
                RsMatrEnsino.Edit
                Status = PgStatusMatricula(RsMatrEnsino.Fields("MatrID"))
                RsMatrEnsino.Fields("StatusMatr") = Status
                RsMatrEnsino.Update
                DoEvents
                pb.Value = pb.Value + 1
                If UCase(Status) = "ATIVO" Then
                        cAtivo = cAtivo + 1
                    Else
                        cInativo = cInativo + 1
                End If
                
                RsMatrEnsino.MoveNext
            Loop
    End If
    RsMatrEnsino.Close
    
    sSQL = "SELECT MatriculaEnsino.MatrID, Matriculas.Nome, Ensino.Descr, MatriculaEnsino.DtInicio, MatriculaEnsino.DtFinal, MatriculaEnsino.StatusMatr " & _
           "FROM (MatriculaEnsino INNER JOIN Matriculas ON MatriculaEnsino.MatrID = Matriculas.MatrID) INNER JOIN Ensino ON MatriculaEnsino.EnsinoID = Ensino.ID " & _
           "WHERE (((MatriculaEnsino.DtInicio) <> Null) AND ((IsNull(MatriculaEnsino.DtFinal)))) " & _
           "ORDER BY MatriculaEnsino.StatusMatr, Ensino.Descr, Matriculas.Nome"
    Call Relatorio(rptListAtivosInativos, sSQL)
    rptListAtivosInativos.Sections("Section2").Controls("lblAtivos").Caption = cAtivo
    rptListAtivosInativos.Sections("Section2").Controls("lblInativos").Caption = cInativo
    rptListAtivosInativos.Show 1

End Sub
