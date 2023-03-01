VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_FiltroProvasAluno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "PGE - Filtro de Provas por Aluno"
   ClientHeight    =   1740
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4710
   Icon            =   "Form_FiltroAlunoConc.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1740
   ScaleWidth      =   4710
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   3015
      TabIndex        =   7
      Top             =   1260
      Width           =   1590
   End
   Begin VB.CommandButton Bt_Aplicar 
      Caption         =   "&Aplicar"
      Height          =   420
      Left            =   3015
      TabIndex        =   6
      Top             =   765
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo:"
      Height          =   1275
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   2865
      Begin MSComCtl2.DTPicker DTP_Inicial 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         Format          =   24510465
         CurrentDate     =   38393
      End
      Begin MSComCtl2.DTPicker DTP_Final 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   780
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         Format          =   24510465
         CurrentDate     =   38393
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Inicial:"
         Height          =   195
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Data Final:"
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   840
         Width           =   795
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "FILTRO DE PROVAS POR ALUNO"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   315
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4680
   End
End
Attribute VB_Name = "Form_FiltroProvasAluno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatrProva As Recordset
Dim QtdProvas As Integer
Dim QtdProvasDia As Integer
'Dim dtAv As String
Dim DtAnt As Date

Private Sub Bt_Aplicar_Click()
    QtdProvas = 0
    QtdProvasDia = 0
    If DTP_Final.Value < DTP_Inicial.Value Then
        MsgBox "A Data Final não deve ser inferior a Data Inicial.", vbInformation, "PGE - Aviso!"
        Exit Sub
    End If
    
    Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE DtAvaliacao >= #" & Format(DTP_Inicial.Value, "MM/DD/YYYY") & "# AND DtAvaliacao <= #" & Format(DTP_Final.Value, "MM/DD/YYYY") & "# ORDER BY DtAvaliacao,MatrID")
    If RsMatrProva.BOF And RsMatrProva.EOF Then
            MsgBox "Nenhuma prova encontrada.", vbInformation, "PGE - Aviso"
            Exit Sub
        Else
            
            RsMatrProva.MoveFirst
            If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
                Exit Sub
            End If
            Call Cab
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
                            Call ImprDados
                        End If
                    Else
                        Call ImprDados
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
Private Sub ImprDados()
    'If Trim(dtAv) = "0" Then Exit Sub
    DoEvents
    ObjPreview.FontSize = CI.tFonte
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Italico
                
    ObjPreview.Print Tab(5); DtAnt; _
                     Tab(30); Left("000", 3 - Len(Trim(QtdProvasDia))) & Trim(QtdProvasDia)
End Sub
Private Sub Cab()
    DoEvents
    Call cPreview(2)
    ObjPreview.FontBold = True
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("RELATÓRIO DE FREQUENCIA DE ALUNOS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value) / 2)
    ObjPreview.Print "RELATÓRIO DE FREQUENCIA DE ALUNOS NO PERIODO DE: " & DTP_Inicial.Value & " ATÉ " & DTP_Final.Value
    ObjPreview.Print
    ObjPreview.Print Tab(5); "Data"; _
                     Tab(25); "Num. de Provas"; ' _
                     Tab(40); "Nome do Aluno"; _
                     Tab(100); "Qtd. provas na data"
End Sub
Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub
Private Function PgTotProvas(m As String, dt As String) As String
    'Erro na quantidade de provas
    Dim RsTMP As Recordset
    Dim Prvs As String
    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & m & "' AND DtAvaliacao = #" & Format(dt, "MM/DD/YYYY") & "#")
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Erro ao localizar Matricula e Qtd de provas", vbInformation, "PGE - Aviso"
            PgTotProvas = "000"
        Else
            RsTMP.MoveLast
            PgTotProvas = RsTMP.RecordCount
    End If
End Function
Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    DTP_Final.Value = Date
    DTP_Inicial.Value = DTP_Final - 30
End Sub
