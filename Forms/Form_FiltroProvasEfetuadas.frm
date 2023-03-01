VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_FiltroProvasEfetuadas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Filtro de Provas Efetuadas"
   ClientHeight    =   5445
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7995
   Icon            =   "Form_FiltroProvasEfetuadas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5445
   ScaleWidth      =   7995
   Begin MSComDlg.CommonDialog cd 
      Left            =   3780
      Top             =   4980
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkExpArquivo 
      Caption         =   "Exportar para arquivo"
      Height          =   195
      Left            =   180
      TabIndex        =   21
      Top             =   5100
      Width           =   3375
   End
   Begin VB.ComboBox cbProf 
      Height          =   315
      Left            =   180
      Style           =   2  'Dropdown List
      TabIndex        =   20
      Top             =   4560
      Width           =   3915
   End
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   6180
      Picture         =   "Form_FiltroProvasEfetuadas.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   4500
      Width           =   1695
   End
   Begin VB.CommandButton Bt_Aplicar 
      Caption         =   "&Aplicar"
      Height          =   735
      Left            =   4440
      Picture         =   "Form_FiltroProvasEfetuadas.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   4500
      Width           =   1695
   End
   Begin VB.Frame Frame4 
      Caption         =   "Critério:"
      Height          =   1275
      Left            =   4020
      TabIndex        =   10
      Top             =   420
      Width           =   3855
      Begin VB.Frame Frame5 
         BorderStyle     =   0  'None
         Height          =   735
         Left            =   2460
         TabIndex        =   22
         Top             =   300
         Width           =   1215
         Begin VB.OptionButton Opt_TipoRpt 
            Caption         =   "Analítico"
            Height          =   195
            Index           =   1
            Left            =   60
            TabIndex        =   24
            Top             =   420
            Visible         =   0   'False
            Width           =   1050
         End
         Begin VB.OptionButton Opt_TipoRpt 
            Caption         =   "Sintético"
            Height          =   195
            Index           =   0
            Left            =   60
            TabIndex        =   23
            Top             =   180
            Value           =   -1  'True
            Visible         =   0   'False
            Width           =   1050
         End
      End
      Begin VB.OptionButton Opt_Criterio 
         Caption         =   "Não Corrigidos"
         Height          =   195
         Index           =   3
         Left            =   240
         TabIndex        =   18
         Top             =   960
         Width           =   1695
      End
      Begin VB.OptionButton Opt_Criterio 
         Caption         =   "Não Habilitado"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   17
         Top             =   720
         Width           =   1695
      End
      Begin VB.OptionButton Opt_Criterio 
         Caption         =   "Habilitado"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   16
         Top             =   480
         Width           =   1395
      End
      Begin VB.OptionButton Opt_Criterio 
         Caption         =   "Todos"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   15
         Top             =   240
         Value           =   -1  'True
         Width           =   1155
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Disciplina:"
      Height          =   2415
      Left            =   4020
      TabIndex        =   5
      Top             =   1800
      Width           =   3855
      Begin VB.CheckBox Chk_DisciplinaTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   3000
         TabIndex        =   9
         Top             =   180
         Width           =   795
      End
      Begin VB.ListBox Lst_Disciplina 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   7
         Top             =   420
         Width           =   3615
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Por Curso:"
      Height          =   2415
      Left            =   60
      TabIndex        =   4
      Top             =   1800
      Width           =   3855
      Begin VB.CheckBox Chk_EnsinoTodos 
         Caption         =   "Todos"
         Height          =   195
         Left            =   3000
         TabIndex        =   8
         Top             =   180
         Width           =   795
      End
      Begin VB.ListBox Lst_Ensino 
         Height          =   1860
         Left            =   120
         Style           =   1  'Checkbox
         TabIndex        =   6
         Top             =   420
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo:"
      Height          =   1275
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   3855
      Begin MSComCtl2.DTPicker DTP_Inicial 
         Height          =   315
         Left            =   1080
         TabIndex        =   11
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         Format          =   55640065
         CurrentDate     =   38393
      End
      Begin MSComCtl2.DTPicker DTP_Final 
         Height          =   315
         Left            =   1080
         TabIndex        =   12
         Top             =   780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         Format          =   55640065
         CurrentDate     =   38393
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Final:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Label Label3 
      Caption         =   "Professor:"
      Height          =   195
      Left            =   180
      TabIndex        =   19
      Top             =   4260
      Width           =   975
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "FILTRO DE PROVAS EFETUADAS"
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
      Width           =   8055
   End
End
Attribute VB_Name = "Form_FiltroProvasEfetuadas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEnsino As Recordset
Dim RsDisciplina As Recordset
Dim RsMatriculaProva As Recordset

Dim Criterio As String
'Dim Criterio2 As String

Dim Aprovado As Integer
Dim lin As Integer

Private Function cEnsino() As String 'Checa os ensinos marcados
    Dim i As Integer
    If Chk_EnsinoTodos.Value = 0 Then
            For i = 0 To Lst_Ensino.ListCount - 1
                If Lst_Ensino.Selected(i) = True Then
                    'cEnsino = IIf(Trim(cEnsino) = "", "Trafego.EnsinoID = " & PgIDEnsino(Lst_Ensino.List(i)), cEnsino & " OR Trafego.EnsinoID = " & PgIDEnsino(Lst_Ensino.List(i)))
                    cEnsino = IIf(Trim(cEnsino) = "", "EnsinoID = " & PgIDEnsino(Lst_Ensino.List(i)), cEnsino & " OR EnsinoID = " & PgIDEnsino(Lst_Ensino.List(i)))
                End If
            Next
            'cEnsino = IIf(Trim(cEnsino) = "", " AND (Trafego.EnsinoID = 0)", " AND (" & cEnsino & ")")
             cEnsino = IIf(Trim(cEnsino) = "", " AND (EnsinoID = 0)", " AND (" & cEnsino & ")")
        Else
            cEnsino = ""
            Exit Function
    End If
End Function
Private Function cDisciplina() As String 'Checa os Disciplinas marcados
    Dim i As Integer
    If Chk_DisciplinaTodos.Value = 0 Then
            For i = 0 To Lst_Disciplina.ListCount - 1
                If Lst_Disciplina.Selected(i) = True Then
                    'cDisciplina = IIf(Trim(cDisciplina) = "", " Trafego.DisciplinaID = " & PgIDDisciplina(Lst_Disciplina.List(i)), cDisciplina & " OR Trafego.DisciplinaID = " & PgIDDisciplina(Lst_Disciplina.List(i)))
                    cDisciplina = IIf(Trim(cDisciplina) = "", " DisciplinaID = " & PgIDDisciplina(Lst_Disciplina.List(i)), cDisciplina & " OR DisciplinaID = " & PgIDDisciplina(Lst_Disciplina.List(i)))
                End If
            Next
            'cDisciplina = IIf(Trim(cDisciplina) = "", " AND (Trafego.DisciplinaID = 0)", " AND (" & cDisciplina & ")")
            cDisciplina = IIf(Trim(cDisciplina) = "", " AND (DisciplinaID = 0)", " AND (" & cDisciplina & ")")
        Else
            cDisciplina = ""
            Exit Function
    End If
End Function

Private Sub Bt_Aplicar_Click()
    Dim TotRegistros As Integer
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    If DTP_Inicial.Value > DTP_Final.Value Then
        MsgBox "A data Inicial não pode ser maior que a final. Por favor verifique.", vbInformation, "CESNet - Atenção"
        Exit Sub
    End If

      'OK================
    'Criterio = "SELECT MatriculaProva.*, Ensino.Descr, Disciplina.Descr " & _
                "FROM (MatriculaProva INNER JOIN Ensino ON MatriculaProva.DisciplinaID = Ensino.ID) INNER JOIN Disciplina ON MatriculaProva.EnsinoID = Disciplina.ID"
     
     Criterio = "SELECT MatriculaProva.*, Ensino.Descr, Disciplina.Descr " & _
                "FROM (MatriculaProva INNER JOIN Ensino ON MatriculaProva.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaProva.DisciplinaID = Disciplina.ID"

                
    Criterio = Criterio & " WHERE DtAvaliacao >= #" & Format(DTP_Inicial.Value, "mm/dd/yyyy") & "# AND DtAvaliacao <= #" & Format(DTP_Final.Value, "mm/dd/yyyy") & "#"
    
    If Trim(cbProf.Text) <> "(Todos)" And Trim(cbProf.Text) <> "" Then
        Criterio = Criterio & " AND ProfIDN = " & Trim(left(cbProf.Text, 3))
    End If
    Select Case Aprovado
        Case 0
        Case 1
            Criterio = Criterio & " AND MatriculaProva.Status = 'HB'"
        Case 2
            Criterio = Criterio & " AND MatriculaProva.Status = 'NH'"
        Case 3
            Criterio = Criterio & " AND MatriculaProva.Status = 'NC'"
    End Select
    
    Criterio = Criterio & cEnsino & cDisciplina & " ORDER BY DtAvaliacao"
    
    Set RsMatriculaProva = BD.OpenRecordset(Criterio)
    If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
            MsgBox "Nenhuma prova foi encontrada.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsMatriculaProva.MoveLast
            TotRegistros = RsMatriculaProva.RecordCount
            RsMatriculaProva.MoveFirst
            'Exportar arquivo======================
            If chkExpArquivo.Value = 1 Then
                Call ExportarArquivo(Criterio, "S")
                Exit Sub
            End If
            '====================================
            
            rptProvasEfetuadas.Sections("cab").Controls("lbProfessor").Caption = IIf(Trim(cbProf.Text) = "", "(Todos)", cbProf.Text)
            rptProvasEfetuadas.Sections("Section2").Controls("lbTotalReg").Caption = TotRegistros
            Call Relatorio(rptProvasEfetuadas, Criterio)
            rptProvasEfetuadas.Show 1

            'If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
            '    Exit Sub
            'End If
           'If Opt_TipoRpt(0) = True Then
           '         Call RptSintetico
           '     Else
           '         Call RptAnalitico
           ' End If
    End If
End Sub
Private Sub RptSintetico()
'Relatorio estatistico de (EnsinoID,DisciplinaID,Horário)
    Dim Estat(20, 50, 2) As Integer
    Dim tempo As Integer
    Dim g1 As Integer
    Dim g2 As Integer
    Dim cont As Integer
    Dim Hora As String
    
    RsMatriculaProva.MoveFirst
    Do Until RsMatriculaProva.EOF
        
        'If Trim(Right(RsMatriculaProva.Fields("DtHrAV"), 8)) < "12:00" Then
        '        tempo = 0 'Manha
        '    Else
        '        tempo = 1 'Tarde
        'End If
        Hora = Trim(Right(RsMatriculaProva.Fields("DtHrAV"), 8))
        If Hora <= "12:00" Then
                tempo = 0
            Else
                If Hora >= "12:01" And Hora <= "18:00" Then
                        tempo = 1
                    Else
                        If Hora >= "18:01" Then
                            tempo = 2
                        End If
                End If
        End If
        
        
        
        
        Estat(RsMatriculaProva.Fields("EnsinoID"), 0, 0) = RsMatriculaProva.Fields("EnsinoID")
        Estat(RsMatriculaProva.Fields("EnsinoID"), RsMatriculaProva.Fields("DisciplinaID"), tempo) = Estat(RsMatriculaProva.Fields("EnsinoID"), RsMatriculaProva.Fields("DisciplinaID"), tempo) + 1
        
        RsMatriculaProva.MoveNext
    Loop
    Call cPreview(1)
    'MANHA
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.FontSize = 12
    ObjPreview.FontBold = True
    ObjPreview.Font = CI.Fonte
    
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO SINTÉTICO DE PROVAS EFETUADAS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value) / 2)
    ObjPreview.Print "RELATÓRIO SINTÉTICO DE PROVAS EFETUADAS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    
    
    ObjPreview.FontSize = CI.tFonte
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(6); "Turno: MANHÃ"
    ObjPreview.FontBold = False
    
    
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Sublinhado
    cont = 0
    
    ObjPreview.Print
    For g1 = 1 To 20
        If Estat(g1, 0, 0) <> 0 Then
            ObjPreview.Print
            ObjPreview.Print Tab(15); "Ensino: " & PgNomeEnsino(g1)
            ObjPreview.Print
        End If
        For g2 = 1 To 50
            If Estat(g1, g2, 0) <> 0 Then
                ObjPreview.Print Tab(20); "Disciplina: " & PgNomeDisciplina(g2); Tab(75); "Qtd. Provas: " & Mid(String(3, "0"), 1, 3 - Len(Trim(Estat(g1, g2, 0)))) & Trim(Estat(g1, g2, 0))
                cont = cont + Val(Estat(g1, g2, 0))
                'ObjPreview.Print
            End If
        Next
    Next
    ObjPreview.Print
    ObjPreview.Print Tab(75); "TOTAL DE PROVAS: " & left("000", 4 - Len(cont)) & cont
    'TARDE
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontSize = CI.tFonte
    
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(6); "Turno: TARDE"
    ObjPreview.FontBold = False
    
    
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Sublinhado
    cont = 0
    ObjPreview.Print
    For g1 = 1 To 20
        If Estat(g1, 0, 0) <> 0 Then
            ObjPreview.Print
            ObjPreview.Print Tab(15); "Ensino: " & PgNomeEnsino(g1)
            ObjPreview.Print
        End If
        For g2 = 1 To 50
            If Estat(g1, g2, 1) <> 0 Then
                ObjPreview.Print Tab(20); "Disciplina: " & PgNomeDisciplina(g2); Tab(75); "Qtd. Provas: " & Mid(String(3, "0"), 1, 3 - Len(Trim(Estat(g1, g2, 1)))) & Trim(Estat(g1, g2, 1))
                cont = cont + Val(Estat(g1, g2, 1))
            End If
        Next
    Next
    ObjPreview.Print
    ObjPreview.Print Tab(75); "TOTAL DE PROVAS: " & left("000", 4 - Len(cont)) & cont
    
    '*********************
    'NOITE
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontSize = CI.tFonte
    
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(6); "Turno: NOITE"
    ObjPreview.FontBold = False
    
    
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Sublinhado
    cont = 0
    ObjPreview.Print
    For g1 = 1 To 20
        If Estat(g1, 0, 0) <> 0 Then
            ObjPreview.Print
            ObjPreview.Print Tab(15); "Ensino: " & PgNomeEnsino(g1)
            ObjPreview.Print
        End If
        For g2 = 1 To 50
            If Estat(g1, g2, 2) <> 0 Then
                ObjPreview.Print Tab(20); "Disciplina: " & PgNomeDisciplina(g2); Tab(75); "Qtd. Provas: " & Mid(String(3, "0"), 1, 3 - Len(Trim(Estat(g1, g2, 2)))) & Trim(Estat(g1, g2, 2))
                cont = cont + Val(Estat(g1, g2, 2))
            End If
        Next
    Next
    ObjPreview.Print
    ObjPreview.Print Tab(75); "TOTAL DE PROVAS: " & left("000", 4 - Len(cont)) & cont

    
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
End Sub
Private Sub RptAnalitico()
    Call Cab
    Dim tmp As String
    lin = 0
    Do Until RsMatriculaProva.EOF
                If RsMatriculaProva.Fields("Aprovado") = True Then
                       tmp = "Aprovado"
                    Else
                        Select Case RsMatriculaProva.Fields("Nota")
                            Case Is >= NotaMedia
                                tmp = "Não Aprovado"
                            Case Is < NotaMedia
                                tmp = "Não Aprovado"
                            Case Else
                                tmp = "Não Corrigido"
                        End Select
                End If
               
                ObjPreview.FontSize = 8 'CI.tFonte
                ObjPreview.Font = CI.Fonte
                ObjPreview.FontBold = CI.Negrito
                ObjPreview.FontItalic = CI.Italico
                ObjPreview.FontUnderline = CI.Sublinhado
                DoEvents
                ObjPreview.Print Tab(6); RsMatriculaProva.Fields("DtAvaliacao"); _
                Tab(19); RsMatriculaProva.Fields("MatrID"); _
                Tab(33); PgDadosMatr(RsMatriculaProva.Fields("MatrID")).Nome; _
                Tab(98); PgNomeEnsino(RsMatriculaProva.Fields("EnsinoID")); _
                Tab(126); PgNomeDisciplina(RsMatriculaProva.Fields("DisciplinaID")); _
                Tab(182); RsMatriculaProva.Fields("NProva"); _
                Tab(189); RsMatriculaProva.Fields("Tipo"); _
                Tab(195); tmp
                'Tab(154); PgNomeModulo(RsMatriculaProva.Fields("ModuloID"));
                lin = lin + 1
                If CI.Preview = True Then
                    If lin >= 46 Then Exit Do
                End If
                RsMatriculaProva.MoveNext
    Loop
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
End Sub
Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub
Private Sub LstEnsino()
    Lst_Ensino.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.EOF Then
            MsgBox "Por favor cheque o cadastro de ENSINO pois nao existe nehum ensino cadastrado", vbInformation, "CESNet - Atenção"
            Bt_Aplicar.Enabled = False
            Exit Sub
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Lst_Ensino.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If
End Sub
Private Sub LstDisciplina()
    Lst_Disciplina.Clear
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina ORDER BY Descr")
    If RsDisciplina.BOF And RsDisciplina.EOF Then
            MsgBox "Por favor cheque o cadastro de Disciplina pois nao existe nehum Disciplina cadastrado", vbInformation, "CESNet - Atenção"
            Bt_Aplicar.Enabled = False
            Exit Sub
        Else
            RsDisciplina.MoveFirst
            Do Until RsDisciplina.EOF
                Lst_Disciplina.AddItem (RsDisciplina.Fields("Descr"))
                RsDisciplina.MoveNext
            Loop
    End If
End Sub




Private Sub cbProf_DropDown()
    Dim RsTMP As Recordset
    cbProf.Clear
    cbProf.AddItem "(Todos)"
    cbProf.Text = "(Todos)"
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Professores ORDER BY Nome")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                cbProf.AddItem left("000", 3 - Len(RsTMP.Fields("ProfID"))) & RsTMP.Fields("ProfID") & " - " & RsTMP.Fields("Nome")
                RsTMP.MoveNext
            Loop
    End If
            
    
End Sub

Private Sub Chk_EnsinoTodos_Click()
     If Chk_EnsinoTodos.Value = 0 Then
            For lin = 0 To Lst_Ensino.ListCount - 1
                Lst_Ensino.Selected(lin) = False
            Next
            Lst_Ensino.Enabled = True
        Else
            For lin = 0 To Lst_Ensino.ListCount - 1
                Lst_Ensino.Selected(lin) = True
            Next
            Lst_Ensino.Enabled = False
    End If
End Sub
Private Sub Chk_DisciplinaTodos_Click()
     If Chk_DisciplinaTodos.Value = 0 Then
            For lin = 0 To Lst_Disciplina.ListCount - 1
                Lst_Disciplina.Selected(lin) = False
            Next
            Lst_Disciplina.Enabled = True
        Else
            For lin = 0 To Lst_Disciplina.ListCount - 1
                Lst_Disciplina.Selected(lin) = True
            Next
            Lst_Disciplina.Enabled = False
    End If
End Sub


Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub


Private Sub ExportarArquivo(sSQL As String, Tp As String)
    Dim RsTMP       As Recordset
    Dim caminho     As String
    Dim sDados      As String
    
    
    Set RsTMP = BD.OpenRecordset(sSQL)
    If RsTMP.BOF And RsTMP.EOF Then
            Exit Sub
        Else
            RsTMP.MoveFirst
    End If
    
    cd.DialogTitle = "Local e Nome do Arquivo?"
    cd.InitDir = App.path
    cd.FileName = "rel_ProvasEfetuadas"
    cd.filter = "Excel |*.xls"
    'cd.Filter = "Todos | *.*"
    
    cd.ShowSave
    caminho = Trim(cd.FileName)
    If caminho = "" Then Exit Sub
    
    Do Until RsTMP.EOF
        'Select Case Tp
        '    Case "S"
        '        sDados = RsTMP.Fields("Nome") & ";" & RsTMP.Fields("DtOrientacao") & ";" & _
        '                RsTMP.Fields("Ensino.Descr") & ";" & RsTMP.Fields("ContarDeDisciplinaID")
        '    Case "A"
        '        sDados = RsTMP.Fields("Professores.Nome") & ";" & RsTMP.Fields("DtOrientacao") & ";" & _
        '                RsTMP.Fields("Ensino.Descr") & ";" & RsTMP.Fields("Disciplina.Descr") & ";" & _
        '                RsTMP.Fields("MatrID") & ";" & RsTMP.Fields("Matriculas.Nome")
        'End Select
                 
        sDados = RsTMP.Fields("DtAvaliacao") & ";" & RsTMP.Fields("MatrID") & ";" & _
                        RsTMP.Fields("Ensino.Descr") & ";" & RsTMP.Fields("Disciplina.Descr") & ";" & _
                        RsTMP.Fields("NProva") & ";" & RsTMP.Fields("Tipo") & ";" & RsTMP.Fields("Status") & ";" & _
                        IIf(Trim(cbProf.Text) = "", "(Todos)", cbProf.Text)
        Call ExpArq(caminho, sDados)
        RsTMP.MoveNext
    Loop
    MsgBox "Arquivo exportado!", vbInformation, "CESNet - Aviso"
End Sub


Private Sub Form_Load()
    DoEvents
    DTP_Inicial.Value = Date
    DTP_Final.Value = Date
    
    LstEnsino
    LstDisciplina
    
    'Chk_CriterioTodos.Value = 1
    Chk_EnsinoTodos.Value = 1
    Chk_DisciplinaTodos.Value = 1
    
End Sub
Private Sub Opt_Criterio_Click(Index As Integer)
    Select Case Index
        Case 0
            Aprovado = 0
        Case 1
            Aprovado = 1
        Case 2
            Aprovado = 2
        Case 3
            Aprovado = 3
    End Select
End Sub
Private Sub Cab()
    Call cPreview(2)
    DoEvents
    ObjPreview.FontSize = 8
    ObjPreview.Font = "Arial"
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(5); "Dt. Avaliação"; _
    Tab(20); "Matricula"; _
    Tab(33); "Nome"; _
    Tab(98); "Ensino"; _
    Tab(126); "Disciplina"; _
    Tab(180); "Prova"; _
    Tab(187); "Tipo"; _
    Tab(195); "Status"
    'Tab(154); "Módulo";
    ObjPreview.FontBold = False
End Sub
