VERSION 5.00
Begin VB.Form Form_AutenticacaoUsuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Autenticação de Usuario"
   ClientHeight    =   2580
   ClientLeft      =   45
   ClientTop       =   360
   ClientWidth     =   6795
   Icon            =   "Form_AutenticacaoUsuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   6795
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   120
      Picture         =   "Form_AutenticacaoUsuario.frx":030A
      ScaleHeight     =   1875
      ScaleWidth      =   1875
      TabIndex        =   5
      Top             =   480
      Width           =   1875
   End
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Height          =   510
      Left            =   4860
      TabIndex        =   4
      Top             =   1080
      Width           =   1770
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "&OK"
      Height          =   510
      Left            =   4860
      TabIndex        =   3
      Top             =   540
      Width           =   1770
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      IMEMode         =   3  'DISABLE
      Left            =   2400
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   1200
      Width           =   2220
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Digite sua senha:"
      Height          =   195
      Left            =   2460
      TabIndex        =   1
      Top             =   900
      Width           =   1275
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "AUTENTICAÇÃO DE USUARIO"
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
      Width           =   6870
   End
End
Attribute VB_Name = "Form_AutenticacaoUsuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Autent As Boolean
Public Function CarregarForm() As Boolean

    Me.Show 1
    CarregarForm = Autent
    Unload Me
End Function

Private Sub Bt_Cancelar_Click()
    Autent = False
    Unload Me
End Sub

Private Sub Bt_OK_Click()
    Chk_Usu
    Unload Me
End Sub

Private Sub Form_Load()
    Autent = False
End Sub
Private Sub Chk_Usu()
    Dim RsUsu As Recordset
    Dim Senha As String
    Set RsUsu = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuarioID)
    If RsUsu.BOF And RsUsu.EOF Then
            MsgBox "Erro ao localizar Usuario", vbInformation, "CESNet - Aviso"
            RsUsu.Close
        Else
            RsUsu.MoveFirst
            Senha = Crypto(RsUsu.Fields("PWD"))
            If Text1.Text = Senha Then
                    Autent = True
                Else
                    MsgBox "Usuário não autenticado!", vbCritical, "CESNet - Aviso"
                    Autent = False
            End If
            RsUsu.Close
    End If
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Bt_OK_Click
End Sub
