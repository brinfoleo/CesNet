VERSION 5.00
Begin VB.Form Form_RelatProvasCad 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Listagem de Provas Cadastradas"
   ClientHeight    =   2235
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6090
   Icon            =   "Form_RelatProvasCad.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2235
   ScaleWidth      =   6090
   Begin VB.Frame Frame1 
      Height          =   1785
      Left            =   45
      TabIndex        =   1
      Top             =   360
      Width           =   5970
      Begin VB.CommandButton Bt_Cancelar 
         Caption         =   "Cancelar"
         Height          =   735
         Left            =   4020
         Picture         =   "Form_RelatProvasCad.frx":030A
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   960
         Width           =   1875
      End
      Begin VB.CommandButton Bt_Imprimir 
         Caption         =   "Imprimir"
         Height          =   735
         Left            =   4020
         Picture         =   "Form_RelatProvasCad.frx":0614
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   180
         Width           =   1875
      End
      Begin VB.ComboBox Cb_Disciplina 
         Height          =   315
         Left            =   1080
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   945
         Width           =   2670
      End
      Begin VB.ComboBox Cb_Ensino 
         Height          =   315
         Left            =   1125
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   405
         Width           =   2625
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Curso:"
         Height          =   195
         Left            =   405
         TabIndex        =   3
         Top             =   450
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Disciplina:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   960
         Width           =   735
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "LISTAGEM DAS PROVAS CADASTRADAS"
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
      Width           =   6225
   End
End
Attribute VB_Name = "Form_RelatProvasCad"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEnsino As Recordset
Dim RsDisciplina As Recordset
Dim RsTrafego As Recordset
Dim RsProva As Recordset

Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub

Private Sub Bt_Imprimir_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    Dim EnsinoID As Integer
    Dim DisciplinaID As Integer
    Dim RefTrafID As Integer
    EnsinoID = PgIDEnsino(Cb_Ensino.Text)
    DisciplinaID = PgIDDisciplina(Cb_Disciplina.Text)
    Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY SerieID,ModuloID")
    If RsTrafego.BOF And RsTrafego.EOF Then
            MsgBox "Nenhuma Referencia de Trafego cadastrada.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsTrafego.MoveFirst
    End If
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, False, False) = False Then
        Exit Sub
    End If
    '****************************
    Call cPreview(4, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Print
    'ObjPreview.FontSize = 16
    'ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("FICHA DE MATRICULA") / 2)
    'ObjPreview.Print "FICHA DE MATRICULA"
    
    'ObjPreview.Print
    ObjPreview.FontSize = CI.tFonte
    'ObjPreview.FontBold = False
    ObjPreview.FontName = CI.Fonte
    'ObjPreview.FontItalic = False
    'ObjPreview.FontUnderline = False
    'ObjPreview.CurrentY = 3000 'vertical
    'ObjPreview.CurrentX = Printer.ScaleWidth - 2500
    'ObjPreview.Print "Matricula:"
    
    'ObjPreview.FontSize = 14
    ObjPreview.FontBold = True
    '*********************
    ObjPreview.Print Tab(5); "Ensino: "; Tab(25); PgNomeEnsino(EnsinoID)
    ObjPreview.Print Tab(5); "Disciplina: "; Tab(25); PgNomeDisciplina(DisciplinaID)
    ObjPreview.FontBold = False
    Do Until RsTrafego.EOF
        ObjPreview.FontBold = True
        ObjPreview.Print
        ObjPreview.Print Tab(10); "Modulo: "; Tab(20); PgNomeModulo(RsTrafego.Fields("ModuloID"))
        ObjPreview.Print Tab(10); "Serie: "; Tab(20); PgNomeSerie(RsTrafego.Fields("SerieID"))
        ObjPreview.FontBold = False
        Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID") & " ORDER BY NProva")
        If RsProva.BOF And RsProva.EOF Then
                ObjPreview.Print Tab(20); "000 - <<< NENHUMA PROVA CADASTRADA >>>"
            Else
                RsProva.MoveFirst
                Do Until RsProva.EOF
                    If Len(RsProva.Fields("Assunto")) > 50 Then
                            ObjPreview.Print Tab(20); RsProva.Fields("NProva") & " -"; _
                                             Tab(26); Mid(RsProva.Fields("Assunto"), 1, 60)
                            ObjPreview.Print Tab(26); Mid(RsProva.Fields("Assunto"), 61); _
                                             Tab(110); "Pag.: " & RsProva.Fields("Pag")
                        Else
                            ObjPreview.Print Tab(20); RsProva.Fields("NProva") & " - "; _
                                             Tab(26); RsProva.Fields("Assunto"); _
                                             Tab(110); "Pag.: " & RsProva.Fields("Pag")
                    End If
                    RsProva.MoveNext
                Loop
        End If
        RsTrafego.MoveNext
    Loop
    Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
End Sub

Private Sub Cb_Disciplina_DropDown()
    Cb_Disciplina.Clear
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & PgIDEnsino(Cb_Ensino.Text))
    If RsDisciplina.BOF And RsDisciplina.BOF Then
            MsgBox "Não existe nenhuma Disciplina cadastrada para este Ensino.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsDisciplina.MoveFirst
            Do Until RsDisciplina.EOF
                Cb_Disciplina.AddItem (PgNomeDisciplina(RsDisciplina.Fields("DisciplinaID")))
                RsDisciplina.MoveNext
            Loop
    End If
End Sub
Private Sub Cb_Ensino_DropDown()
    Cb_Ensino.Clear
    Cb_Disciplina.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino") ' WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND IsNull(DtConclusao)")
    If RsEnsino.BOF And RsEnsino.BOF Then
            MsgBox "Não existe nenhuma Ensino cadastrado.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_Ensino.AddItem (PgNomeEnsino(RsEnsino.Fields("ID")))
                RsEnsino.MoveNext
            Loop
    End If

End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

