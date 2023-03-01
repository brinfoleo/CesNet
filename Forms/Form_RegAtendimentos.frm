VERSION 5.00
Begin VB.Form Form_RegAtendimentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Registro de Atendimentos Diversos"
   ClientHeight    =   3150
   ClientLeft      =   5715
   ClientTop       =   5415
   ClientWidth     =   7140
   Icon            =   "Form_RegAtendimentos.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3150
   ScaleWidth      =   7140
   Begin VB.TextBox txtHr 
      Height          =   285
      Left            =   1200
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   960
      Width           =   2475
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   4140
      Top             =   1140
   End
   Begin VB.CommandButton btoGravar 
      Caption         =   "&Gravar"
      Height          =   555
      Left            =   4800
      TabIndex        =   7
      Top             =   480
      Width           =   2175
   End
   Begin VB.TextBox txtMotivo 
      Height          =   975
      Left            =   60
      MaxLength       =   64000
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   6
      Text            =   "Form_RegAtendimentos.frx":030A
      Top             =   2040
      Width           =   6975
   End
   Begin VB.TextBox txtUsuario 
      Height          =   285
      Left            =   1260
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   1320
      Width           =   2475
   End
   Begin VB.TextBox txtDt 
      Height          =   285
      Left            =   1200
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   600
      Width           =   2475
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Hora:"
      Height          =   195
      Left            =   120
      TabIndex        =   9
      Top             =   1020
      Width           =   915
   End
   Begin VB.Label Label3 
      Caption         =   "Motivo:"
      Height          =   195
      Left            =   120
      TabIndex        =   3
      Top             =   1740
      Width           =   1035
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Usuario(a):"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   855
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Data:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   660
      Width           =   915
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "REGISTRO DE ATENDIMENTOS DIVERSOS"
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
      Width           =   7170
   End
End
Attribute VB_Name = "Form_RegAtendimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btoGravar_Click()
    Dim Rst As Recordset
    If ValidarSoftware("RegAtendimento") = False Then Exit Sub
    Set Rst = BD.OpenRecordset("SELECT * FROM RegAtendimento")
    Rst.AddNew
    Rst.Fields("UsuID") = Trim(left(txtUsuario.Text, 4))
    Rst.Fields("Dt") = txtDt.Text
    Rst.Fields("Hr") = txtHr.Text
    Rst.Fields("Motivo") = txtMotivo.Text
    Rst.Update
    MsgBox "Atendimento registrado com sucesso!", vbInformation, "Aviso"
    LimpForm
End Sub

Private Sub Form_Load()
    LimpForm
    txtUsuario.Text = left("0000", 4 - Len(Trim(UsuarioID))) & UsuarioID & " - " & Usuario
End Sub
Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub LimpForm()
    txtDt.Text = Date
    txtHr.Text = Time
    txtUsuario.Text = left("0000", 4 - Len(Trim(UsuarioID))) & UsuarioID & " - " & Usuario
    txtMotivo.Text = ""
End Sub
Private Sub HDForm(op As Boolean)
    txtDt.Enabled = op
    txtHr.Enabled = op
    txtUsuario.Enabled = op
    txtMotivo.Enabled = op
End Sub

Private Sub Timer1_Timer()
    txtHr.Text = Time
End Sub

Private Sub txtHr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtDt_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub txtUsuario_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
