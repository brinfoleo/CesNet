VERSION 5.00
Begin VB.Form Form_MailEnviar 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Envio de Mail"
   ClientHeight    =   6150
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   10920
   Icon            =   "Form_MailEnviar.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6150
   ScaleWidth      =   10920
   Begin VB.TextBox txtAssunto 
      Height          =   315
      Left            =   1140
      MaxLength       =   200
      TabIndex        =   2
      Top             =   900
      Width           =   9615
   End
   Begin VB.CommandButton btEnviar 
      Caption         =   "&Enviar"
      Height          =   915
      Left            =   8700
      Picture         =   "Form_MailEnviar.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5220
      Width           =   2115
   End
   Begin VB.TextBox txtDescr 
      Height          =   3735
      Left            =   120
      MaxLength       =   64000
      MultiLine       =   -1  'True
      TabIndex        =   3
      Top             =   1380
      Width           =   10635
   End
   Begin VB.TextBox txtPara 
      Height          =   315
      Left            =   1140
      TabIndex        =   1
      Top             =   480
      Width           =   6015
   End
   Begin VB.CommandButton btPara 
      Caption         =   "&Para..."
      Height          =   375
      Left            =   180
      TabIndex        =   0
      Top             =   480
      Width           =   915
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Assunto:"
      Height          =   195
      Left            =   240
      TabIndex        =   6
      Top             =   960
      Width           =   735
   End
   Begin VB.Label Label22 
      Alignment       =   2  'Center
      BackColor       =   &H00000080&
      Caption         =   "NOVA MENSAGEM"
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
      TabIndex        =   5
      Top             =   0
      Width           =   10935
   End
End
Attribute VB_Name = "Form_MailEnviar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btEnviar_Click()
    Dim RsMail      As Recordset
    Dim Destino     As String
    Dim EndDest     As String
    
    If Trim(txtAssunto.Text) = "" Then
        MsgBox "O campo ASSUNTO não pode ser em branco!", vbInformation, "CESNet - Aviso"
        txtAssunto.SetFocus
        Exit Sub
    End If
        
    HDForm (False)
    Set RsMail = BD.OpenRecordset("SELECT * FROM Mail")
    
    EndDest = Trim(txtPara.Text)
    
    Destino = Mid(EndDest, InStr(EndDest, "(") + 1, InStr(EndDest, ")") - 2)
    EndDest = Mid(EndDest, InStr(EndDest, ")") + 1, Len(EndDest))
    
    Do Until Trim(Destino) = ""
        RsMail.AddNew
        RsMail.Fields("Data") = Date
        RsMail.Fields("DE") = UsuarioID
        RsMail.Fields("PARA") = PgIDUsuResp(Destino)
        RsMail.Fields("Assunto") = Trim(txtAssunto.Text)
        RsMail.Fields("Descr") = IIf(Trim(txtDescr.Text) = "", Null, txtDescr.Text)
        RsMail.Fields("Novo") = True
        RsMail.Update
        If Trim(EndDest) = "" Then
                Destino = ""
            Else
                Destino = Mid(EndDest, InStr(EndDest, "(") + 1, InStr(EndDest, ")") - 2)
                EndDest = Mid(EndDest, InStr(EndDest, ")") + 1, Len(EndDest))
        End If
    Loop
    HDForm (True)
    MsgBox "Mensagem enviada com sucesso!", vbInformation, "CESNet - Aviso"
    Unload Me
End Sub

Private Sub btPara_Click()
   txtPara.Text = Form_MailContatos.PgUsuContatos
End Sub

Private Sub txtPara_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub HDForm(op As Boolean)
    btPara.Enabled = op
    btEnviar.Enabled = op
    txtAssunto.Enabled = op
    txtDescr.Enabled = op
End Sub

Private Function PgIDUsuResp(txtResp As String) As Integer
    Dim RsTMP As Recordset
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Usuario WHERE Responsavel = '" & txtResp & "'")
    If RsTMP.BOF And RsTMP.EOF Then
            PgIDUsuResp = 0
        Else
            RsTMP.MoveFirst
            PgIDUsuResp = RsTMP.Fields("UsuarioID")
    End If
    RsTMP.Close
End Function

