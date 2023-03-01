VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_SQLExecute 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - SQL Execute"
   ClientHeight    =   6165
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9225
   Icon            =   "Form_SQLExecute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6165
   ScaleWidth      =   9225
   Begin VB.Frame Frame1 
      Height          =   5685
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   9105
      Begin VB.CommandButton Bt_Executar 
         Caption         =   "Executar"
         Enabled         =   0   'False
         Height          =   510
         Left            =   6660
         TabIndex        =   7
         Top             =   945
         Width           =   2265
      End
      Begin VB.Frame Frame3 
         Caption         =   "Arquivo:"
         Height          =   645
         Left            =   135
         TabIndex        =   5
         Top             =   180
         Width           =   8835
         Begin VB.Label Lb_Arq 
            Caption         =   "..."
            Height          =   285
            Left            =   180
            TabIndex        =   6
            Top             =   225
            Width           =   8520
         End
      End
      Begin VB.CommandButton Bt_PgArq 
         Caption         =   "Buscar Arquivo"
         Height          =   510
         Left            =   4365
         TabIndex        =   4
         Top             =   945
         Width           =   2265
      End
      Begin VB.Frame Frame2 
         Height          =   4110
         Left            =   135
         TabIndex        =   2
         Top             =   1440
         Width           =   8835
         Begin VB.ListBox Lst_Status 
            Height          =   3765
            Left            =   135
            TabIndex        =   3
            Top             =   225
            Width           =   8565
         End
      End
      Begin MSComDlg.CommonDialog CD_Conexao 
         Left            =   3465
         Top             =   990
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "SQL EXECUTE"
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
      Width           =   9360
   End
End
Attribute VB_Name = "Form_SQLExecute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim caminho As String
Dim Cmd(100) As String
Dim cont        As Integer
Private Sub PgDados(iCaminho As String)
    On Error GoTo TratErro
    Dim Arquivo     As String
   
    cont = 1
    Arquivo = FreeFile
    Open iCaminho For Input As Arquivo
    Do
        
        Line Input #Arquivo, Cmd(cont)
        If Trim(Cmd(cont)) = "" Then Exit Do
        cont = cont + 1
    Loop
    
    
    Close #Arquivo
    Exit Sub
TratErro:
    If Err.Number = 62 Then
            If cont <> 0 Then cont = cont - 1
            Exit Sub
        Else
            MsgBox Err.Description, vbInformation, Err.Number
    End If
End Sub

Private Sub Bt_Executar_Click()
    Lst_Status.Clear
    ExecutarSQL
End Sub

Private Sub Bt_PgArq_Click()
  With CD_Conexao
        .DialogTitle = "CESNet - SQL Execute"
        .InitDir = App.path
        .filter = "SQL Execute|*.sql"
        .DefaultExt = "*.sql"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        If Trim(.FileName) = "" Then Exit Sub
        If Len(.FileName) >= 200 Then
                MsgBox "Nome do local para o arquivo é muito extenso. Por favor modifique!", vbInformation, "CESNet - Aviso!"
                Bt_Executar.Enabled = False
            Else
                caminho = Trim(.FileName)
                Lb_Arq.Caption = caminho
                PgDados (caminho)
                Lst_Status.Clear
                Lst_Status.AddItem "==============================================================="
                Lst_Status.AddItem " Número de Comandos: " & left(String(3, "0"), 3 - Len(Trim(cont))) & cont
                Lst_Status.AddItem "==============================================================="
                Lst_Status.AddItem " "
                Bt_Executar.Enabled = True
        
        End If
        
    End With

End Sub
Private Sub ExecutarSQL()

    On Error GoTo TratSQL
    Dim ErrDescr As String
    Call RegLog("0", "MODIFICANDO BASE DE DADOS COM ARQUIVO: " & caminho)
    Dim xCont As Integer
    xCont = 1
    Lst_Status.AddItem " "
    Lst_Status.AddItem "****************** INICIANDO DO PROCESSO ******************"
    Lst_Status.AddItem " "
    Do Until xCont > cont
        Lst_Status.AddItem left(String(3, "0"), 3 - Len(Trim(xCont))) & xCont & " - [" & Now & "] comando: " & Cmd(xCont)
        'Lst_Status.AddItem Cmd(xCont)
        BD.Execute Cmd(xCont)
        xCont = xCont + 1
    Loop
    Lst_Status.AddItem " "
    Lst_Status.AddItem "********************* FIM DO PROCESSO *************************"
    Exit Sub
TratSQL:
    ErrDescr = Trim(Err.Number) & " - " & Trim(Err.Description)
    Call RegLog(Err.Number, "ERRO EXECUTE SQL: " & ErrDescr)
    Lst_Status.AddItem "        ERRO: " & ErrDescr
    'MsgBox Err.Number
    Lst_Status.AddItem " "
    Resume Next
    
End Sub



Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

