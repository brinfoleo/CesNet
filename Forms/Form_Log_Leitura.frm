VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_Log_Leitura 
   Caption         =   "CESNET - Leitura de Log"
   ClientHeight    =   4470
   ClientLeft      =   3675
   ClientTop       =   5280
   ClientWidth     =   13035
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   4470
   ScaleWidth      =   13035
   Begin MSComDlg.CommonDialog cmdLog 
      Left            =   12300
      Top             =   420
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.TreeView trvLog 
      Height          =   3255
      Left            =   120
      TabIndex        =   4
      Top             =   960
      Width           =   12675
      _ExtentX        =   22357
      _ExtentY        =   5741
      _Version        =   393217
      Style           =   7
      Appearance      =   1
   End
   Begin VB.Frame Frame1 
      Height          =   795
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   12375
      Begin VB.CommandButton btoAbrir 
         Caption         =   "&Abrir"
         Height          =   435
         Left            =   11460
         TabIndex        =   3
         Top             =   240
         Width           =   855
      End
      Begin VB.CommandButton btoBuscar 
         Caption         =   "&Buscar"
         Height          =   435
         Left            =   10620
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
      Begin VB.TextBox txtArquivo 
         Height          =   315
         Left            =   120
         TabIndex        =   1
         Top             =   300
         Width           =   10395
      End
   End
End
Attribute VB_Name = "Form_Log_Leitura"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cont As Integer

Private Sub btoAbrir_Click()
    On Error GoTo SaiComando
    If Trim(txtArquivo.Text) = "" Then Exit Sub
    cont = 0
    
    trvLog.Nodes.Clear
    Dim ArqLog      As String
    Dim texto       As String
    ArqLog = FreeFile
    Open txtArquivo.Text For Input As ArqLog
    Line Input #ArqLog, texto
    Do Until Trim(texto) = ""
        
        PreencherArvore Mid(texto, 2, 10), Mid(texto, 12, Len(texto)), texto
        Line Input #ArqLog, texto
    Loop
    Close #ArqLog
    Exit Sub

SaiComando:
    Exit Sub
    
End Sub

Private Sub btoBuscar_Click()
    cmdLog.DialogTitle = "Buscar registro de log"
    cmdLog.FileName = "*.txt"
    'cmdLog.filter = "*.log"
    cmdLog.ShowOpen
    txtArquivo.Text = cmdLog.FileName
End Sub
Private Function PreencherArvore(Grupo As String, subgrupo As String, sTexto As String)
    On Error Resume Next
    cont = cont + 1
    trvLog.Nodes.Add , , Grupo, "Dt.:" & Grupo
    trvLog.Nodes.Add Grupo, tvwChild, "sub" & cont, sTexto

'    TrVw_Usu.Nodes.Add , , GrupoID, NGrupo, 1
'    TrVw_Usu.Nodes.Add GrupoID, tvwChild, subGrupoID, nSubGrupo, 2
End Function
Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Resize()
    trvLog.Width = Form_Log_Leitura.Width - 400
    
    trvLog.Height = Form_Log_Leitura.Height - (Frame1.Height + 1000)
    
End Sub
