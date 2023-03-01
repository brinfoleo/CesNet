VERSION 5.00
Begin VB.Form Form_TrocaPwd 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Trocar Senha"
   ClientHeight    =   3135
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6195
   Icon            =   "Form_TrocaPwd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3135
   ScaleWidth      =   6195
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2715
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   6075
      Begin VB.Frame Frame2 
         Height          =   1515
         Left            =   3480
         TabIndex        =   12
         Top             =   1140
         Width           =   2535
         Begin VB.CommandButton Bt_Cancelar 
            Caption         =   "&Cancelar"
            Height          =   555
            Left            =   120
            TabIndex        =   14
            Top             =   840
            Width           =   2295
         End
         Begin VB.CommandButton Bt_GrvSenha 
            Caption         =   "&Gravar"
            Height          =   555
            Left            =   135
            TabIndex        =   13
            Top             =   180
            Width           =   2295
         End
      End
      Begin VB.TextBox Txt_OldPwd 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   10
         Top             =   2160
         Width           =   1935
      End
      Begin VB.TextBox Txt_NovaPwd2 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   9
         Top             =   1680
         Width           =   1935
      End
      Begin VB.TextBox Txt_NovaPwd1 
         Appearance      =   0  'Flat
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1380
         MaxLength       =   6
         PasswordChar    =   "*"
         TabIndex        =   8
         Top             =   1200
         Width           =   1935
      End
      Begin VB.Label Lb_Resp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1350
         TabIndex        =   7
         Top             =   675
         Width           =   4575
      End
      Begin VB.Label Lb_Chave 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Left            =   1380
         TabIndex        =   6
         Top             =   240
         Width           =   2415
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsavel:"
         Height          =   195
         Left            =   360
         TabIndex        =   5
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Login:"
         Height          =   195
         Left            =   840
         TabIndex        =   4
         Top             =   255
         Width           =   495
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha Antiga:"
         Height          =   195
         Left            =   300
         TabIndex        =   3
         Top             =   2220
         Width           =   1035
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Confirmar Senha:"
         Height          =   195
         Left            =   120
         TabIndex        =   2
         Top             =   1740
         Width           =   1215
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nova Senha:"
         Height          =   195
         Left            =   360
         TabIndex        =   1
         Top             =   1260
         Width           =   975
      End
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "TROCAR SENHA"
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
      Left            =   45
      TabIndex        =   11
      Top             =   0
      Width           =   6135
   End
End
Attribute VB_Name = "Form_TrocaPwd"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUsuario As Recordset
Dim OldPwd As String
Dim NewPwd As String
Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub

Private Sub Bt_GrvSenha_Click()
    AbrirBD_DAO
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuarioID)
    If RsUsuario.BOF And RsUsuario.EOF Then
        MsgBox "Erro ao localizar Usuario. Chame o suporte.", vbInformation, "CESNet - Troca Senha"
        Exit Sub
    End If
    OldPwd = Crypto(RsUsuario.Fields("Pwd"))
    If Txt_OldPwd.Text = OldPwd Then
            If Txt_NovaPwd1.Text = Txt_NovaPwd2.Text Then
                    If Txt_NovaPwd1.Text = Txt_OldPwd.Text Then
                            MsgBox "Sua nova senha não pode ser igual a sua senha antiga!", vbInformation, "CESNet - Troca Senha"
                            Exit Sub
                        Else
                            RsUsuario.Edit
                            RsUsuario.Fields("Pwd") = Crypto(Txt_NovaPwd1.Text)
                            RsUsuario.Fields("TrocarPwd") = False
                            RsUsuario.Fields("Acesso") = 0
                            RsUsuario.Update
                            BD.Close
                            MsgBox "Senha Alterada com sucesso" & Chr(13) & "Por favor reinicie o sistema.", vbInformation, "CESNet - Troca Senha"
                            Unload Me
                            End
                    End If
                Else
                    MsgBox "Nova senha invalida!", vbInformation, "CESNet - Troca Senha"
                    Exit Sub
            End If
        Else
            MsgBox "Senha Invalida!", vbInformation, "CESNet - Troca Senha"
    End If
End Sub
Public Function TrocarSenha(UsuID As Integer)
    UsuarioID = UsuID
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuID)
    If RsUsuario.BOF And RsUsuario.EOF Then
            MsgBox "Não foi possivel localizar o Usuario. Por favor tente novamente!", vbInformation, "CESNet - Aviso"
            Exit Function
        Else
            RsUsuario.MoveFirst
            Lb_Chave.Caption = RsUsuario.Fields("Chv")
            Lb_Resp.Caption = RsUsuario.Fields("Responsavel")
    End If
    Form_TrocaPwd.Show 1
End Function
