VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_ConectarBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Conexão a Base de Dados"
   ClientHeight    =   4815
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5850
   Icon            =   "Form_ConectarBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   4815
   ScaleWidth      =   5850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   2760
      Left            =   60
      TabIndex        =   1
      Top             =   2040
      Width           =   5730
      Begin VB.Frame Fr_Cliente 
         Caption         =   "Cliente:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   135
         TabIndex        =   5
         Top             =   615
         Visible         =   0   'False
         Width           =   5505
         Begin VB.TextBox Txt_IP_Serv 
            Height          =   330
            Left            =   1350
            MaxLength       =   15
            TabIndex        =   8
            Top             =   540
            Width           =   2850
         End
         Begin VB.CommandButton Bt_ConectarServ 
            Caption         =   "&Conectar"
            Height          =   780
            Left            =   3420
            Picture         =   "Form_ConectarBD.frx":030A
            Style           =   1  'Graphical
            TabIndex        =   7
            Top             =   1140
            Width           =   1995
         End
         Begin VB.TextBox Txt_Serv_Porta 
            Height          =   285
            Left            =   1350
            MaxLength       =   4
            TabIndex        =   6
            Text            =   "4444"
            Top             =   1215
            Width           =   690
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "IP do Servidor:"
            Height          =   195
            Left            =   135
            TabIndex        =   11
            Top             =   585
            Width           =   1050
         End
         Begin VB.Label Label4 
            Caption         =   "Exemplo: 192.168.0.1"
            ForeColor       =   &H00808080&
            Height          =   195
            Left            =   1350
            TabIndex        =   10
            Top             =   900
            Width           =   1680
         End
         Begin VB.Label Label9 
            Caption         =   "Porta:"
            Height          =   240
            Left            =   720
            TabIndex        =   9
            Top             =   1215
            Width           =   420
         End
      End
      Begin VB.Frame Fr_Servidor 
         Caption         =   "Servidor:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1995
         Left            =   135
         TabIndex        =   4
         Top             =   615
         Visible         =   0   'False
         Width           =   5505
         Begin VB.Frame Fr_ServLocal 
            Caption         =   "Local:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   135
            TabIndex        =   20
            Top             =   450
            Visible         =   0   'False
            Width           =   5235
            Begin MSComDlg.CommonDialog CD_Conexao 
               Left            =   180
               Top             =   810
               _ExtentX        =   847
               _ExtentY        =   847
               _Version        =   393216
            End
            Begin VB.CommandButton Bt_ServLocalLoc 
               Caption         =   "..."
               BeginProperty Font 
                  Name            =   "Arial Black"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   900
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   240
               Left            =   4680
               TabIndex        =   23
               Top             =   315
               Width           =   330
            End
            Begin VB.CommandButton Bt_ServLocalGravar 
               Caption         =   "&Gravar"
               Height          =   720
               Left            =   3285
               Picture         =   "Form_ConectarBD.frx":0614
               Style           =   1  'Graphical
               TabIndex        =   22
               Top             =   630
               Width           =   1860
            End
            Begin VB.TextBox Txt_ServLocal 
               Height          =   285
               Left            =   180
               TabIndex        =   21
               Top             =   315
               Width           =   4290
            End
         End
         Begin VB.OptionButton Opt_Serv 
            Caption         =   "Local"
            Height          =   195
            Index           =   1
            Left            =   2880
            TabIndex        =   19
            Top             =   225
            Width           =   690
         End
         Begin VB.OptionButton Opt_Serv 
            Caption         =   "Rede"
            Height          =   195
            Index           =   0
            Left            =   1530
            TabIndex        =   18
            Top             =   225
            Width           =   690
         End
         Begin VB.Frame Fr_ServRede 
            Caption         =   "Rede:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1410
            Left            =   135
            TabIndex        =   12
            Top             =   450
            Visible         =   0   'False
            Width           =   5235
            Begin VB.TextBox Txt_IP_Serv_Local 
               Enabled         =   0   'False
               Height          =   285
               Left            =   90
               MaxLength       =   15
               TabIndex        =   15
               Top             =   450
               Width           =   2850
            End
            Begin VB.CommandButton Bt_AtivServ 
               Caption         =   "Ativar Servidor"
               Height          =   375
               Left            =   2295
               TabIndex        =   14
               Top             =   810
               Width           =   2130
            End
            Begin VB.TextBox Txt_Serv_Porta_Local 
               Height          =   285
               Left            =   3375
               MaxLength       =   4
               TabIndex        =   13
               Text            =   "4444"
               Top             =   450
               Width           =   690
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "IP do Servidor:"
               Height          =   195
               Left            =   90
               TabIndex        =   17
               Top             =   225
               Width           =   1050
            End
            Begin VB.Label Label8 
               Caption         =   "Porta:"
               Height          =   240
               Left            =   3375
               TabIndex        =   16
               Top             =   225
               Width           =   510
            End
         End
      End
      Begin VB.ComboBox Cb_SN 
         Height          =   315
         ItemData        =   "Form_ConectarBD.frx":091E
         Left            =   1935
         List            =   "Form_ConectarBD.frx":0928
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   270
         Width           =   1365
      End
      Begin VB.Label Label1 
         Caption         =   "Você esta no Servidor?"
         Height          =   240
         Left            =   135
         TabIndex        =   2
         Top             =   315
         Width           =   1725
      End
   End
   Begin VB.Image Image1 
      Height          =   1710
      Left            =   0
      Picture         =   "Form_ConectarBD.frx":0936
      Stretch         =   -1  'True
      Top             =   0
      Width           =   5895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CONEXÃO A BASE DE DADOS"
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
      Top             =   1680
      Width           =   5880
   End
End
Attribute VB_Name = "Form_ConectarBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Banco As Database
Dim RsLocal As Recordset
Option Explicit


Private Sub Bt_AtivServ_Click()
 Dim Arquivo As String
    On Error GoTo ErroLocate
    
    If Trim(Txt_IP_Serv_Local.Text) = "" Then
        MsgBox "O número de IP não pode ser um valor nulo!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
     If Trim(Txt_Serv_Porta_Local.Text) = "" Then
        MsgBox "O número da PORTA não pode ser um valor nulo!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    
    Arquivo = FreeFile
    Open App.Path & "\CESNet.srv" For Output As Arquivo
    
    'Print #Arquivo, "[SERVIDOR]"
    Print #Arquivo, "1" '- um: É o servidor
    'Print #Arquivo, "[IP]"
    Print #Arquivo, Txt_IP_Serv_Local.Text
    'Print #Arquivo, "[PORTA]"
    Print #Arquivo, Txt_Serv_Porta_Local.Text
    
    
    Close #Arquivo
    MsgBox "Favor reiniciar o Aplicativo", vbInformation, "CESNet"
    End
    Exit Sub
ErroLocate:
    Call RegLogErros(Err.Number, Err.Description, "Form_ConectarBD", UsuarioID)
    MsgBox Err.Description, vbInformation, Err.Number
    End

End Sub

Private Sub Bt_ConectarServ_Click()

    Dim Arquivo As String
    On Error GoTo ErroLocate
    
    If Trim(Txt_IP_Serv.Text) = "" Then
        MsgBox "O número de IP não pode ser um valor nulo!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
     If Trim(Txt_Serv_Porta.Text) = "" Then
        MsgBox "O número da PORTA não pode ser um valor nulo!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    
    Arquivo = FreeFile
    Open App.Path & "\CESNet.srv" For Output As Arquivo
    
    'Print #Arquivo, "[SERVIDOR]"
    Print #Arquivo, "0" '- zero: Não é o servidor
    'Print #Arquivo, "[IP]"
    Print #Arquivo, Txt_IP_Serv.Text
    'Print #Arquivo, "[PORTA]"
    Print #Arquivo, Txt_Serv_Porta.Text
    
    'Print #Arquivo, Crypto("[SERVIDOR]"))
    'Print #Arquivo, "0" '- zero: Não é o servidor
    'Print #Arquivo, "[IP]"
    'Print #Arquivo, Txt_IP_Serv.Text
    'Print #Arquivo, "[PORTA]"
    'Print #Arquivo, Txt_Serv_Porta.Text
    
    Close #Arquivo
    MsgBox "Favor reiniciar o Aplicativo", vbInformation, "CESNet"
    End
    Exit Sub
ErroLocate:
    Call RegLogErros(Err.Number, Err.Description, "Form_ConectarBD", UsuarioID)
    MsgBox Err.Description, vbInformation, Err.Number
    End

    
End Sub

Private Sub Bt_ServLocalGravar_Click()

    Dim Arquivo As String
    On Error GoTo ErroLocate
    
    'If Trim(Txt_IP_Serv.Text) = "" Then
    '    MsgBox "O número de IP não pode ser um valor nulo!", vbInformation, "CESNet - Aviso!"
    '    Exit Sub
    'End If
     If Trim(Txt_ServLocal.Text) = "" Then
        MsgBox "O LOCAL não pode ser um valor nulo!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    
    Arquivo = FreeFile
    Open App.Path & "\CESNet.srv" For Output As Arquivo
    
    'Print #Arquivo, "[SERVIDOR]"
    Print #Arquivo, "2" '- 2: Endereço local
    'Print #Arquivo, "[IP]"
    Print #Arquivo, Trim(Txt_ServLocal.Text)
    'Print #Arquivo, "[PORTA]"
    Print #Arquivo, "0" 'Txt_Serv_Porta.Text
    
    
    Close #Arquivo
    MsgBox "Favor reiniciar o aplicativo CESNet.", vbInformation, "CESNet"
    End
    Exit Sub
ErroLocate:
    Call RegLogErros(Err.Number, Err.Description, "Form_ConectarBD", UsuarioID)
    MsgBox Err.Description, vbInformation, Err.Number
    End

    

End Sub


Private Sub Bt_ServLocalLoc_Click()
    Dim tmpPath As String
    With CD_Conexao
        .DialogTitle = "CESNet - Conexão ao Banco de Dados"
        .InitDir = App.Path
        .Filter = "Banco de Dados|*.mdb"
        .filename = "Dados.mdb"
        .DefaultExt = "*.mdb"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        .filename = Left(.filename, Len(.filename) - Len(.FileTitle) - 1)
        If Len(.filename) >= 200 Then
                MsgBox "Nome do local para Banco de Dados muito extenso. Por favor modifique", vbInformation, "CesNet - Aviso!"
                Txt_ServLocal.SetFocus
            Else
                
                tmpPath = IIf(IsNull(PathBD), PathBD, .filename)
                tmpPath = Mid(.filename, 1, InStr(.filename, "Database") - 2)
                Txt_ServLocal.Text = tmpPath
        End If
    End With

End Sub

Private Sub Cb_SN_Click()
    Select Case Cb_SN.Text
        Case "Sim"
            Fr_Servidor.Visible = True
            Fr_Cliente.Visible = False
        Case "Não"
            Fr_Servidor.Visible = False
            Fr_Cliente.Visible = True
        Case Else
            Fr_Servidor.Visible = False
            Fr_Cliente.Visible = False
    End Select
End Sub

Private Sub Form_Load()
    Dim a, b, c As Integer
    Dim VerSis As String
    a = App.Major
    b = App.Minor
    c = App.Revision
    VerSis = Mid("00", 1, 2 - Len(a)) & a & "." & Mid("00", 1, 2 - Len(b)) & b & "." & Mid("000", 1, 3 - Len(Trim(c))) & Trim(c)
    Me.Caption = "CESNet - Conexão ao Servidor [v." & VerSis & "]"
    
    Txt_IP_Serv_Local.Text = MDIForm_Main.Winsock_Main.LocalIP

    
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub



Private Sub Txt_LocLog_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub


Private Sub Opt_Serv_Click(Index As Integer)
    Select Case Index
        Case 0
            Fr_ServRede.Visible = True
            Fr_ServLocal.Visible = False
        Case 1
            Fr_ServRede.Visible = False
            Fr_ServLocal.Visible = True
        Case Else
            Fr_ServRede.Visible = False
            Fr_ServLocal.Visible = False
        End Select
End Sub
