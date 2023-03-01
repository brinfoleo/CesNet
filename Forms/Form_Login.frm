VERSION 5.00
Begin VB.Form Form_Login 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "CESNet - Login"
   ClientHeight    =   4590
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7755
   DrawMode        =   6  'Mask Pen Not
   DrawStyle       =   1  'Dash
   Icon            =   "Form_Login.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   PaletteMode     =   1  'UseZOrder
   Picture         =   "Form_Login.frx":030A
   ScaleHeight     =   4590
   ScaleWidth      =   7755
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Txt_Chv 
      Appearance      =   0  'Flat
      Height          =   285
      Left            =   2715
      MaxLength       =   30
      TabIndex        =   4
      Top             =   2760
      Width           =   3855
   End
   Begin VB.TextBox Txt_Pwd 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      IMEMode         =   3  'DISABLE
      Left            =   2715
      MaxLength       =   30
      PasswordChar    =   "*"
      TabIndex        =   6
      Top             =   3120
      Width           =   3855
   End
   Begin VB.CommandButton Bt_Ok 
      Caption         =   "&Ok"
      Height          =   675
      Left            =   3300
      Picture         =   "Form_Login.frx":4F6A
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3660
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Canc 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   675
      Left            =   5040
      Picture         =   "Form_Login.frx":5274
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   3660
      Width           =   1575
   End
   Begin VB.Label Label4 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1035
      Left            =   2040
      TabIndex        =   0
      Top             =   1080
      Width           =   2190
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Nome do Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Left            =   720
      TabIndex        =   7
      Top             =   2805
      Width           =   1740
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Senha:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   1065
      TabIndex        =   5
      Top             =   3165
      Width           =   1395
   End
   Begin VB.Label Label3 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CES"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   540
      TabIndex        =   3
      Top             =   360
      Width           =   2610
   End
   Begin VB.Label Label7 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicativo de Gerenciamento para o CES"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3480
      TabIndex        =   2
      Top             =   1140
      Width           =   3645
   End
   Begin VB.Label lbVersaoAno 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   645
      Left            =   4320
      TabIndex        =   1
      Top             =   1680
      Width           =   3300
   End
End
Attribute VB_Name = "Form_Login"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Dim RsUsuario As Recordset
'Dim RsMaxAcessos As Recordset
Dim PwdSys As String
Dim tmp As Integer
Dim Sit As Boolean
Private Sub Bt_Canc_Click()
    Sit = False
    Unload Me
End Sub

Private Sub Bt_OK_Click()
    If Trim(Txt_Chv.Text) = "" Or Trim(Txt_Pwd.Text) = "" Then
        MsgBox "CHAVE ou SENHA inválida", vbCritical, "CESNet - Login"
        Exit Sub
    End If
   ValidaAcesso
End Sub


Private Sub Form_Load()
    Form_Login.Caption = "CESNet - Login  [Versão: " & Versao & "]"
    lbVersaoAno.Caption = VersaoAno
End Sub





'Private Sub Label8_Click()
'    Txt_Chv.Text = "LEO"
'    Txt_Pwd.Text = "leo"
'    Bt_Ok_Click
'End Sub

Private Sub Txt_Chv_KeyPress(KeyAscii As Integer)
   KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 0 Then Exit Sub
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    If KeyAscii >= 65 And KeyAscii <= 90 Then 'Letras Maiusculas
            Else
                If KeyAscii >= 48 And KeyAscii <= 57 Then 'Numeros entre 0 a 9
                        Exit Sub
                    Else
                        KeyAscii = 0
                        Beep
                End If
    End If
End Sub
Private Sub ValidaAcesso()
    Dim RsUsuario As Recordset
    
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE Chv = '" & Trim(Txt_Chv.Text) & "'")

    If RsUsuario.BOF And RsUsuario.EOF Then
            MsgBox "CHAVE ou SENHA invalida, por favor verifique a tecla CAPS LOCK e tente novamente!", vbInformation, "CESNet - Login"
            Exit Sub
        Else
        PwdSys = Crypto(RsUsuario.Fields("Pwd"))
            UsuarioID = RsUsuario.Fields("UsuarioID")
            Usuario = Trim(Txt_Chv.Text)
            If Trim(Txt_Pwd.Text) = PwdSys Then
                    
                        If RsUsuario.Fields("TrocarPWD") = True Then
                                MsgBox "Sua senha deve ser trocada agora.", vbInformation, "CESNet - Login"
                                Form_TrocaPwd.TrocarSenha (UsuarioID)
                                'Unload Me
                                Exit Sub
                            Else
                                RsUsuario.Edit
                                RsUsuario.Fields("Acesso") = Val(IIf(IsNull(RsUsuario.Fields("Acesso")), 0, RsUsuario.Fields("Acesso"))) + 1
                                If Val(RsUsuario.Fields("Acesso")) + 1 >= MaxAcessos Then
                                    If RsUsuario.Fields("PWDNuncaExp") = False Then
                                        RsUsuario.Fields("TrocarPwd") = True
                                    End If
                                End If
                                RsUsuario.Update
                        End If
                    

                    Sit = True
                    LoadGrupoUsu (RsUsuario.Fields("GrupoID"))
                    RsUsuario.Close
                    Unload Me
                    
                Else
                    MsgBox "CHAVE ou SENHA invalida, por favor verifique a tecla CAPS LOCK e tente novamente!", vbInformation, "CESNet - Login"

            End If
    End If
End Sub

Private Sub Txt_Pwd_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Bt_OK_Click
        KeyAscii = 0
    End If
End Sub
Public Function LoadLogin() As Double
    Sit = False
    Form_Login.Show 1
    LoadLogin = Sit
End Function
