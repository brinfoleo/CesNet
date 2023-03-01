VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_FiltroProvasAluno 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Filtro de Provas Aplicadas por Aluno"
   ClientHeight    =   2760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6285
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2760
   ScaleWidth      =   6285
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   420
      Left            =   4620
      TabIndex        =   7
      Top             =   1200
      Width           =   1590
   End
   Begin VB.CommandButton Bt_Aplicar 
      Caption         =   "&Aplicar"
      Height          =   420
      Left            =   4620
      TabIndex        =   6
      Top             =   720
      Width           =   1590
   End
   Begin VB.Frame Frame1 
      Height          =   2175
      Left            =   45
      TabIndex        =   1
      Top             =   405
      Width           =   4365
      Begin VB.ComboBox cbUsuario 
         Height          =   315
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1500
         Width           =   3855
      End
      Begin MSComCtl2.DTPicker DTP_Inicial 
         Height          =   315
         Left            =   1080
         TabIndex        =   2
         Top             =   300
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         Format          =   16384001
         CurrentDate     =   38393
      End
      Begin MSComCtl2.DTPicker DTP_Final 
         Height          =   315
         Left            =   1080
         TabIndex        =   3
         Top             =   780
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         CalendarTitleBackColor=   16711680
         CalendarTitleForeColor=   16777215
         Format          =   16384001
         CurrentDate     =   38393
      End
      Begin VB.Label Label3 
         Caption         =   "Usuário:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   1260
         Width           =   855
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
      BackColor       =   &H00C00000&
      Caption         =   "FILTRO DE PROVAS APLICADAS POR ALUNO"
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
      Width           =   6300
   End
End
Attribute VB_Name = "Form_FiltroProvasAluno"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Bt_Aplicar_Click()
    Dim sSQL    As String
    Dim RsTMP   As Recordset
    Dim TotAval As Integer
    sSQL = "SELECT MatriculaProva.DtAvaliacao, MatriculaProva.MatrID, Matriculas.Nome, Ensino.Descr, Disciplina.Descr, MatriculaProva.NProva, MatriculaProva.Assunto, MatriculaProva.Tipo, MatriculaProva.Status, Usuario.Responsavel " & _
         "FROM Usuario INNER JOIN (((Matriculas INNER JOIN MatriculaProva ON Matriculas.MatrID = MatriculaProva.MatrID) INNER JOIN Ensino ON MatriculaProva.EnsinoID = Ensino.ID) INNER JOIN Disciplina ON MatriculaProva.DisciplinaID = Disciplina.ID) ON Usuario.UsuarioID = MatriculaProva.UsuarioIDAv " & _
         "WHERE (((MatriculaProva.DtAvaliacao)>=#" & Format(DTP_Inicial.Value, "mm/dd/yyyy") & "# And (MatriculaProva.DtAvaliacao)<=#" & Format(DTP_Final.Value, "mm/dd/yyyy") & "#))"
         If cbUsuario.Text = "" Or cbUsuario.Text = "(Todos)" Then
            Else
                sSQL = sSQL & " AND UsuarioIDAv = " & CInt(left(cbUsuario.Text, 3))
         End If
         
         sSQL = sSQL & " ORDER BY DtAvaliacao, UsuarioIDAV"

   
         
         
         
    Set RsTMP = BD.OpenRecordset(sSQL)
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Nenhuma prova foi encontrada.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsTMP.MoveLast
            TotAval = RsTMP.RecordCount
            'RsMatriculaProva.MoveFirst
            'rptProvasEfetuadas.Sections("cab").Controls("lbProfessor").Caption = IIf(Trim(cbProf.Text) = "", "(Todos)", cbProf.Text)
            rptListProvasAplicadas.Sections("Section2").Controls("lbTotal").Caption = TotAval
            Call Relatorio(rptListProvasAplicadas, sSQL)
            rptListProvasAplicadas.Show 1
    End If
    RsTMP.Close

End Sub

Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub

Private Sub cbUsuario_DropDown()
    cbUsuario.Clear
    Dim RsUsu As Recordset
    Set RsUsu = BD.OpenRecordset("SELECT * FROM Usuario ORDER BY Responsavel")
    If RsUsu.BOF And RsUsu.EOF Then
        Else
            RsUsu.MoveFirst
            cbUsuario.AddItem "(Todos)"
            Do Until RsUsu.EOF
                cbUsuario.AddItem left("000", 3 - Len(Trim(RsUsu.Fields("UsuarioID")))) & RsUsu.Fields("UsuarioID") & _
                                  " - " & RsUsu.Fields("Responsavel")
                RsUsu.MoveNext
            Loop
            'cbUsuario.Text = cbUsuario.List(0)
            RsUsu.Close
    End If
    
End Sub


Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    DTP_Final.Value = Date
    DTP_Inicial.Value = DTP_Final - 30
End Sub
