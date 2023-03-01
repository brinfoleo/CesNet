VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form Form_Splash 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   Caption         =   "CESNet - Splash"
   ClientHeight    =   6450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11235
   Icon            =   "Form_Splash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "Form_Splash.frx":030A
   ScaleHeight     =   6450
   ScaleWidth      =   11235
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ProgressBar PB_Load 
      Height          =   195
      Left            =   225
      TabIndex        =   10
      Top             =   5490
      Width           =   10545
      _ExtentX        =   18600
      _ExtentY        =   344
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Min             =   1
      Scrolling       =   1
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   360
      Top             =   1755
   End
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   10260
      Top             =   3660
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label5 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Net"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1575
      Left            =   2400
      TabIndex        =   13
      Top             =   1080
      Width           =   3435
   End
   Begin VB.Label lbVersaoAno 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "0000"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   525
      Left            =   5940
      TabIndex        =   12
      Top             =   1980
      Width           =   4920
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Carregando..."
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   315
      TabIndex        =   11
      Top             =   5265
      Width           =   960
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "CES"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1635
      Left            =   600
      TabIndex        =   9
      Top             =   120
      Width           =   3615
   End
   Begin VB.Shape Shape1 
      BackColor       =   &H00FFFFFF&
      BorderWidth     =   5
      Height          =   6225
      Left            =   120
      Top             =   120
      Width           =   10995
   End
   Begin VB.Label Label7 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Aplicativo de Gerenciamento para o CEJA"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   195
      Left            =   6540
      TabIndex        =   8
      Top             =   180
      Width           =   4500
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Versão:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   9660
      TabIndex        =   7
      Top             =   1440
      Width           =   600
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackColor       =   &H80000014&
      BackStyle       =   0  'Transparent
      Caption         =   "Service Pack:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   9270
      TabIndex        =   6
      Top             =   1695
      Width           =   1005
   End
   Begin VB.Label Label6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"Form_Splash.frx":7BF0
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   450
      Left            =   180
      TabIndex        =   5
      Top             =   5820
      Width           =   10860
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "Licenciado para:"
      ForeColor       =   &H00C0C0C0&
      Height          =   195
      Left            =   330
      TabIndex        =   4
      Top             =   4095
      Width           =   1185
   End
   Begin VB.Label Lb_ProdID 
      Alignment       =   1  'Right Justify
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "0000000000000"
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
      Height          =   195
      Left            =   4440
      TabIndex        =   3
      Top             =   5100
      Width           =   6555
   End
   Begin VB.Label Lb_Versao 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "00.00"
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
      Height          =   195
      Left            =   10305
      TabIndex        =   2
      Top             =   1470
      Width           =   555
   End
   Begin VB.Label Lb_SP 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "000"
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
      Height          =   195
      Left            =   10305
      TabIndex        =   1
      Top             =   1695
      Width           =   555
   End
   Begin VB.Label Lb_Licenciado 
      BackColor       =   &H00FFFFFF&
      BackStyle       =   0  'Transparent
      Caption         =   "xxxxxxxxxxxxxx"
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
      Height          =   240
      Left            =   330
      TabIndex        =   0
      Top             =   4365
      Width           =   5685
   End
End
Attribute VB_Name = "Form_Splash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim StatusEnvio As Boolean
Dim txtSubject As String
Dim txtMsg As String



Private Sub cmdEnviar()
  'Verificar se nenhuma conexão está em andamento
    If Winsock1.Tag = "" Then
        If Winsock1.State <> sckClosed Then Winsock1.Close
        Winsock1.Connect "smtp.gmail.com", 25
    End If
    '***************************
    txtSubject = "CESNet - v." & Versao & " (" & Now & ")"
    txtMsg = UnidadeEnsinoNome & vbCrLf & "PathServ: " & PathServ & vbCrLf & "PathBD: " & PathBD & vbCrLf & "IP Serve: " & Winsock1.LocalIP
    If StatusEnvio = True Then
            'ReceberDadosExternos = True
        Else
            'ReceberDadosExternos = False
    End If

End Sub


Private Sub Form_Load()
    PB_Load.Min = 0
    PB_Load.Max = 25
    Lb_Versao.Caption = left(Versao, Len(Versao) - 4)
    Lb_SP.Caption = Right(Versao, 3)
    lbVersaoAno.Caption = VersaoAno
    Lb_Licenciado.Caption = UnidadeEnsinoNome
    Lb_ProdID.Caption = "ProdutoID: " & SoftwareID
    
End Sub

Private Sub Timer1_Timer()
    If CInt(PB_Load.Value) = PB_Load.Max Then
        Unload Me
    End If
    PB_Load.Value = CInt(PB_Load.Value) + 1
    
End Sub

Private Sub Winsock1_Connect()
    On Error Resume Next
    Winsock1.Tag = "conectado"
    Me.Hide
End Sub

Private Sub Winsock1_DataArrival(ByVal bytesTotal As Long)

Dim strData  As String
Dim MsgTexto As String
Dim msg      As String
Dim Status   As String
Dim Erro     As Boolean

If Trim(Winsock1.Tag) <> "" Then
  Winsock1.GetData strData
  Status = left(strData, 3)
  
  'Verifica de o servidor retornou alguma msg de erro
  Select Case Status
     Case "250", "220", "354", "221", "334", "235": Erro = False
     Case Else:
       Erro = True
       Winsock1.Tag = "fechar"
       Status = Mid(strData, 4)
  End Select
  
  Select Case Winsock1.Tag
    Case "conectado":
      'If PgDadosConfig.MailAutenticacao = 1 Then ' chkAuth Then
        msg = "ehlo " & Winsock1.LocalIP & vbCrLf
        Winsock1.Tag = "autenticar"
      'Else
      '  msg = "helo " & Winsock1.LocalIP & vbCrLf
      '  Winsock1.Tag = "conectou"
      'End If
      
      Winsock1.SendData msg
      'stbConexao.Panels(1).Text = "Conectado."
    
    Case "autenticar":
      msg = "auth login" & vbCrLf
      Winsock1.SendData msg
      Winsock1.Tag = "autenticar_usuario"
    
    Case "autenticar_usuario":
      msg = sBase64Encode("eletrosoft.suporte@gmail.com") & vbCrLf
      Winsock1.SendData msg
      Winsock1.Tag = "autenticar_senha"
    
    Case "autenticar_senha":
      msg = sBase64Encode("k3bw8200") & vbCrLf
      Winsock1.SendData msg
      Winsock1.Tag = "conectou"

    Case "conectou":
      'stbConexao.Panels(1).Text = "Enviando..."
      Winsock1.SendData "mail from:<" & Trim("eletrosoft.suporte@gmail.com") & ">" & vbCrLf
      Winsock1.Tag = "from"
    
    Case "from":
      Winsock1.SendData "rcpt to:<" & Trim("eletrosoft.suporte@gmail.com") & ">" & vbCrLf
      
      'Com copia ***********************************
        'If PgDadosConfig.MailRecCopia = 1 Then
        '    Winsock1.Tag = "to"
        '    Winsock1.SendData "rcpt to:<" & Trim(PgDadosConfig.MailEndereco) & ">" & vbCrLf
        'End If
      '*****************************************************
      Winsock1.Tag = "to"
    
    Case "to":
      Winsock1.SendData "data" & vbCrLf
      Winsock1.Tag = "data"
      
    Case "data":
      'A sequencia "." e quebra de linha deve ser substituida por ".." e quebra de linha
      'para evitar que o servidor entenda fim de email antes do fim do texto
      MsgTexto = txtMsg & vbCrLf
      While InStr(MsgTexto, vbCrLf & "." & vbCrLf) <> 0
        MsgTexto = Replace(MsgTexto, vbCrLf & "." & vbCrLf, vbCrLf & ".." & vbCrLf)
      Wend
      
      msg = "subject: " & txtSubject & vbCrLf
      'Mensagem em HTML
      'If chkHTML = vbChecked Then
      '  Msg = Msg & "MIME-Version: 1.0" & vbCrLf & "Content-type: text/html; charset=iso-8859-1" & vbCrLf
      'End If
      msg = msg & MsgTexto & vbCrLf & "." & vbCrLf
      
      'msg = msg & UUEncodeFile(Trim(txtAnexo.Text)) & vbCrLf & "." & vbCrLf
      
      Winsock1.SendData msg
      Winsock1.Tag = "fim"
      
    Case "fim":
      'stbConexao.Panels(1).Text = "Desconectando..."
      Winsock1.SendData "quit" & vbCrLf
      Winsock1.Tag = "fechar"
      
    Case "fechar":
      If Not Erro Then
        'stbConexao.Panels(1).Text = "Enviado com sucesso!"
        StatusEnvio = True
      Else
        'stbConexao.Panels(1).Text = "Erro ao enviar email!"
        'MsgBox Status, vbCritical, "Erro"
        StatusEnvio = False
      End If
      
      Winsock1.Close
      Winsock1.Tag = ""
      'Unload Me
  End Select
  
End If

End Sub

Private Sub Winsock1_Error(ByVal Number As Integer, Description As String, ByVal Scode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
  'MsgBox "Erro ao conectar" & vbNewLine & "Verifique sua conexão ou o endereço do servidor", vbCritical, "Erro"
End Sub

