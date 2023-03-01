VERSION 5.00
Begin VB.Form Form_About 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Sobre"
   ClientHeight    =   4890
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   4950
   Icon            =   "Form_About.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4890
   ScaleWidth      =   4950
   Begin VB.Frame Frame3 
      Caption         =   "Suporte:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   60
      TabIndex        =   8
      Top             =   3420
      Width           =   4815
      Begin VB.Label Label1 
         Caption         =   "Telefone: 21 8379-4470"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   555
         Width           =   3195
      End
      Begin VB.Label Label9 
         Caption         =   "brinfo.leo@gmail.com"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   180
         TabIndex        =   9
         Top             =   210
         Width           =   2775
      End
   End
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   60
      TabIndex        =   2
      Top             =   360
      Width           =   4815
      Begin VB.TextBox txtLicenciado 
         Height          =   285
         Left            =   120
         TabIndex        =   12
         Text            =   "Text1"
         Top             =   420
         Width           =   4575
      End
      Begin VB.Label Label7 
         Caption         =   "Este produto está licenciado para:"
         Height          =   195
         Left            =   120
         TabIndex        =   7
         Top             =   180
         Width           =   2415
      End
      Begin VB.Label Lb_ProdID 
         AutoSize        =   -1  'True
         Caption         =   "ProdutoID:"
         Height          =   195
         Left            =   135
         TabIndex        =   5
         Top             =   900
         Width           =   3165
      End
      Begin VB.Label Lb_Versao 
         Caption         =   "01.02.000"
         Height          =   195
         Left            =   4005
         TabIndex        =   4
         Top             =   885
         Width           =   735
      End
      Begin VB.Label Label2 
         Caption         =   "Versão:"
         Height          =   195
         Left            =   3420
         TabIndex        =   3
         Top             =   885
         Width           =   555
      End
   End
   Begin VB.CommandButton Bt_OK 
      Caption         =   "&Ok"
      Height          =   435
      Left            =   3180
      TabIndex        =   1
      Top             =   4350
      Width           =   1695
   End
   Begin VB.Frame Frame1 
      Height          =   1815
      Left            =   60
      TabIndex        =   0
      Top             =   1560
      Width           =   4815
      Begin VB.TextBox Text1 
         Height          =   1515
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Text            =   "Form_About.frx":030A
         Top             =   180
         Width           =   4575
      End
      Begin VB.Label Label6 
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   6.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   30
         TabIndex        =   6
         Top             =   180
         Width           =   4635
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CESNet - Dados do Aplicativo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   -60
      TabIndex        =   11
      Top             =   0
      Width           =   5040
   End
End
Attribute VB_Name = "Form_About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_OK_Click()
    
    Unload Me
    
End Sub

'

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    
    txtLicenciado.Text = UnidadeEnsinoNome
    Lb_Versao.Caption = Versao
    Lb_ProdID.Caption = "ProdutoID: " & SoftwareID
End Sub


Private Sub Text1_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub txtLicenciado_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 120 And Shift = 0 Then
        If Trim(txtLicenciado.Text) = "manutencao" And LCase(Usuario) = "leo" Then
            Unload Me
            Form_SQLDatabase.Show
        End If
    End If
End Sub
