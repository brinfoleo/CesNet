VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_MatriculaAvisoPreview 
   BackColor       =   &H00000000&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Mensagem"
   ClientHeight    =   8580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   12210
   ForeColor       =   &H00000000&
   Icon            =   "Form_MatriculaAvisoPreview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8580
   ScaleWidth      =   12210
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   450
      Left            =   8160
      Top             =   2100
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Height          =   4875
      Left            =   240
      TabIndex        =   0
      Top             =   3120
      Width           =   11295
      Begin VB.TextBox Txt_Nome 
         Height          =   285
         Left            =   1380
         TabIndex        =   4
         Top             =   360
         Width           =   6555
      End
      Begin VB.TextBox Txt_Texto 
         Height          =   1575
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   3000
         Width           =   10875
      End
      Begin MSFlexGridLib.MSFlexGrid msfgAvisos 
         Height          =   1635
         Left            =   180
         TabIndex        =   8
         Top             =   900
         Width           =   10935
         _ExtentX        =   19288
         _ExtentY        =   2884
         _Version        =   393216
         Cols            =   6
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_MatriculaAvisoPreview.frx":030A
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Aluno:"
         Height          =   195
         Left            =   780
         TabIndex        =   3
         Top             =   420
         Width           =   495
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Texto:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   2760
         Width           =   495
      End
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "Clique na parte preta para fechar!"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   5580
      TabIndex        =   7
      Top             =   8160
      Width           =   6315
   End
   Begin VB.Label Label5 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "ATENÇÃO"
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   72
         Charset         =   0
         Weight          =   900
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   1695
      Left            =   60
      TabIndex        =   6
      Top             =   120
      Width           =   11715
   End
   Begin VB.Image imgNegado 
      Height          =   2550
      Left            =   8100
      Picture         =   "Form_MatriculaAvisoPreview.frx":03AE
      Stretch         =   -1  'True
      Top             =   0
      Width           =   3150
   End
   Begin VB.Label Lb_Bloq 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00000000&
      Caption         =   "MATRÍCULA BLOQUEADA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   5940
      TabIndex        =   5
      Top             =   2520
      Visible         =   0   'False
      Width           =   5715
   End
End
Attribute VB_Name = "Form_MatriculaAvisoPreview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim bloq As Boolean

Private Sub Form_Click()
    Unload Me
End Sub

Private Sub imgNegado_Click()
    Unload Me
End Sub

Private Sub Label5_Click()
    Unload Me
End Sub

Private Sub Label6_Click()
    Unload Me
End Sub

Private Sub Lb_Bloq_Click()
    Unload Me
End Sub

Private Sub msfgAvisos_Click()
    With msfgAvisos
        'RegID = .TextMatrix(.Row, 0)
        'DTP_Dt.Value = .TextMatrix(.Row, 1)
        Txt_Texto.Text = .TextMatrix(.Row, 3)
    
        'Chk_Avisar.Value = IIf(.TextMatrix(.Row, 4) = "SIM", 1, 0)
        'Chk_Bloquear.Value = IIf(Trim(.TextMatrix(.Row, 5)) <> "", 1, 0)
                
        'DTP_Bloqueio.Value = IIf(Trim(.TextMatrix(.Row, 5)) <> "", .TextMatrix(.Row, 5), Date)
    End With
End Sub
Private Sub Timer1_Timer()
    If bloq = False Then Exit Sub
    If Lb_Bloq.Visible = True Then
            Lb_Bloq.Visible = False
            imgNegado.Visible = False
        Else
            Lb_Bloq.Visible = True
            imgNegado.Visible = True
    End If
End Sub



Private Sub Txt_Nome_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Txt_Texto_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
    
End Sub
Public Sub CarregarForm(MatrID As String, Bloquear As Boolean)
    Dim RsAviso As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "' " & _
           "AND dtAvisar<= #" & Format(Date, "MM/DD/YYYY") & "#"
    
    
    'Txt_DtInclusao.Text = DtInclusao
    'Txt_Texto.Text = texto
    'Txt_DtBloqueio.Text = DtBloqueio
    Txt_Nome.Text = PgDadosMatr(MatrID).Nome
    'lbCadastro.Caption = "Cadastrado por " & strUsu
    msfgAvisos.Rows = 1
    Set RsAviso = BD.OpenRecordset(sSQL)
    If RsAviso.BOF And RsAviso.EOF Then
            'LimpForm
        Else
            RsAviso.MoveFirst
            With msfgAvisos
                Do Until RsAviso.EOF
                    DoEvents
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = RsAviso.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = RsAviso.Fields("DtInclusao")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsAviso.Fields("Codigo")), " ", RsAviso.Fields("Codigo"))
                    .TextMatrix(.Rows - 1, 3) = RsAviso.Fields("texto")
                    .TextMatrix(.Rows - 1, 4) = IIf(RsAviso.Fields("Avisar") = True, IIf(IsNull(RsAviso.Fields("DtAvisar")), Date, RsAviso.Fields("DtAvisar")), " ")
                    .TextMatrix(.Rows - 1, 5) = IIf(IsNull(RsAviso.Fields("DtBloqueio")), " ", RsAviso.Fields("DtBloqueio"))
                    RsAviso.MoveNext
                   
                Loop
            End With
            'Bt_Gravar.Enabled = op
            'Bt_Excluir.Enabled = op
    End If
    bloq = Bloquear
    
    
    If Bloquear = True Then
            Lb_Bloq.Visible = True
            imgNegado.Visible = True
        Else
            imgNegado.Visible = False
    End If
    Form_MatriculaAvisoPreview.Show 1
    
End Sub
