VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_BiblFiltro 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Filtro de Livro"
   ClientHeight    =   6075
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10815
   Icon            =   "Form_BiblFiltro.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6075
   ScaleWidth      =   10815
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   8040
      Picture         =   "Form_BiblFiltro.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   5220
      Width           =   2625
   End
   Begin VB.Frame Frame3 
      Height          =   780
      Left            =   7740
      TabIndex        =   12
      Top             =   4185
      Width           =   2985
      Begin VB.ComboBox Cb_Organizar 
         Height          =   315
         ItemData        =   "Form_BiblFiltro.frx":0614
         Left            =   405
         List            =   "Form_BiblFiltro.frx":0624
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   270
         Width           =   2445
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Ordernar por:"
         Height          =   195
         Left            =   90
         TabIndex        =   13
         Top             =   0
         Width           =   960
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Filtro:"
      Height          =   1770
      Left            =   90
      TabIndex        =   3
      Top             =   4185
      Width           =   7530
      Begin VB.TextBox Txt_Assunto 
         Height          =   285
         Left            =   990
         TabIndex        =   11
         Top             =   1350
         Width           =   5730
      End
      Begin VB.ComboBox Cb_Disciplina 
         Height          =   315
         Left            =   990
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   945
         Width           =   4650
      End
      Begin VB.TextBox Txt_Autor 
         Height          =   285
         Left            =   990
         TabIndex        =   9
         Top             =   585
         Width           =   5685
      End
      Begin VB.TextBox Txt_Titulo 
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   225
         Width           =   5685
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Assunto:"
         Height          =   195
         Left            =   225
         TabIndex        =   7
         Top             =   1440
         Width           =   690
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Disciplina:"
         Height          =   195
         Left            =   135
         TabIndex        =   6
         Top             =   1035
         Width           =   780
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Autor:"
         Height          =   195
         Left            =   450
         TabIndex        =   5
         Top             =   630
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Titulo:"
         Height          =   195
         Left            =   450
         TabIndex        =   4
         Top             =   315
         Width           =   465
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Livro(s):"
      Height          =   3615
      Left            =   90
      TabIndex        =   1
      Top             =   405
      Width           =   10635
      Begin MSFlexGridLib.MSFlexGrid MSFG_Livros 
         Height          =   3255
         Left            =   180
         TabIndex        =   2
         Top             =   225
         Width           =   10320
         _ExtentX        =   18203
         _ExtentY        =   5741
         _Version        =   393216
         Cols            =   5
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_BiblFiltro.frx":064C
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "FILTRO DE LIVRO"
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
      Width           =   11010
   End
End
Attribute VB_Name = "Form_BiblFiltro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsBiblInd       As Recordset

Dim RsDiscipl       As Recordset

Dim LivroID         As String
Dim Filtro          As String
Dim LivroSel        As Integer


Private Sub Bt_Cancelar_Click()
    LivroSel = 0
    Unload Me
End Sub

Private Sub Cb_Disciplina_DropDown()
    Cb_Disciplina.Clear
    Set RsDiscipl = BD.OpenRecordset("SELECT * FROM Disciplina ORDER BY Descr")
    If RsDiscipl.BOF And RsDiscipl.EOF Then
            
            Exit Sub
        Else
            RsDiscipl.MoveFirst
            Cb_Disciplina.AddItem ("(Todas)")
            Do Until RsDiscipl.EOF
                Cb_Disciplina.AddItem (RsDiscipl.Fields("Descr"))
                RsDiscipl.MoveNext
            Loop
    End If
End Sub

Private Sub Cb_Organizar_Click()
    FiltrarLivros
End Sub

Private Sub MSFG_Livros_DblClick()
    With MSFG_Livros
        If .Rows = 1 Then Exit Sub
        LivroSel = .TextMatrix(.Row, 0)
        Unload Me
    End With
End Sub

Private Sub Txt_Assunto_Change()
    FiltrarLivros
End Sub
Private Sub Txt_Autor_Change()
    FiltrarLivros
End Sub
Private Sub Txt_Titulo_Change()
    FiltrarLivros
End Sub
Private Sub Cb_Disciplina_Click()
    FiltrarLivros
End Sub
Private Sub Form_Load()
    Cb_Organizar.Text = "LivroID"
End Sub

Private Sub FiltrarLivros()
    Dim ordem As String
    Dim MsgTmp As String
    Select Case Cb_Organizar.Text
        Case "Disciplina"
            ordem = " ORDER BY " & Cb_Organizar.Text & "ID ASC"
        Case Else
            ordem = " ORDER BY " & Cb_Organizar.Text & " ASC"
    End Select
    Filtro = ""
    Filtro = " Titulo LIKE '*" & Trim(Txt_Titulo.Text) & "*'"
    Filtro = Filtro & " AND Autor LIKE '*" & Trim(Txt_Autor.Text) & "*'"
    If Cb_Disciplina.Text = "(Todas)" Then
        Else
            Filtro = IIf(PgIDDisciplina(Cb_Disciplina.Text) = 0, Filtro, Filtro & " AND DisciplinaID = " & PgIDDisciplina(Cb_Disciplina.Text))
    End If
    Filtro = Filtro & " AND Assunto LIKE '*" & Trim(Txt_Assunto.Text) & "*'"
    '*********************************************
    Set RsBiblInd = BD.OpenRecordset("SELECT * FROM BibliotecaIndice WHERE " & Filtro & ordem)
    If RsBiblInd.BOF And RsBiblInd.EOF Then
            MSFG_Livros.Rows = 1
            MSFG_Livros.Rows = 2
            Exit Sub
        Else
            With RsBiblInd
                .MoveFirst
                MSFG_Livros.Rows = 1
                Do Until .EOF
                    DoEvents
                    MSFG_Livros.Rows = MSFG_Livros.Rows + 1
                    LivroID = Mid(String(4, "0"), 1, 4 - Len(.Fields("LivroID"))) & .Fields("LivroID")
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 0) = LivroID
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 1) = .Fields("Titulo")
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 2) = .Fields("Autor")
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 3) = PgNomeDisciplina(.Fields("DisciplinaID"))
                    
                    If .Fields("Emprestado") = True Then
                        'MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 4) = "SIM"
                        MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 4) = IIf(IsNull(.Fields("DtEmprestimo")), "", .Fields("DtEmprestimo"))
                        MSFG_Livros.Col = 1
                        MSFG_Livros.Row = MSFG_Livros.Rows - 1
                        MSFG_Livros.ColSel = MSFG_Livros.Cols - 1
                        MSFG_Livros.FillStyle = flexFillRepeat
                        MSFG_Livros.CellForeColor = &HFF&
                    End If
                    
                    'Set RsBiblAss = BD.OpenRecordset("SELECT * FROM BibliotecaAssunto WHERE LivroID = " & LivroID & " AND Assunto LIKE '*" & Trim(Txt_Assunto.Text) & "*'")
                    'If RsBiblAss.BOF And RsBiblAss.EOF Then
                    '        MsgTmp = "<NENHUM ASSUNTO CADASTRADO>"
                    '    Else
                    '        RsBiblAss.MoveFirst
                    '        Do Until RsBiblAss.EOF
                    '            MsgTmp = IIf(Trim(MsgTmp) = "", Trim(UCase(RsBiblAss.Fields("Assunto"))), MsgTmp & "; " & Trim(UCase(RsBiblAss.Fields("Assunto"))))
                    '            RsBiblAss.MoveNext
                    '        Loop
                    'End If
                    MSFG_Livros.TextMatrix(MSFG_Livros.Rows - 1, 4) = IIf(Trim(.Fields("Assunto")) = "", " ", .Fields("Assunto"))
                    MSFG_Livros.Col = 4
                    MSFG_Livros.Row = MSFG_Livros.Rows - 1
                    MSFG_Livros.ColSel = 4
                    MSFG_Livros.RowSel = MSFG_Livros.Rows - 1
                    MSFG_Livros.FillStyle = flexFillRepeat
                    MSFG_Livros.CellAlignment = 1
                    .MoveNext
                Loop
            End With
    End If
End Sub


Public Function CarregarFormulario() As Integer
    Form_BiblFiltro.Show 1
    CarregarFormulario = LivroSel
End Function


