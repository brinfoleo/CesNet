VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_InfoBD 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Informações do Banco de Dados"
   ClientHeight    =   2310
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8775
   Icon            =   "Form_InfoBD.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2310
   ScaleWidth      =   8775
   Begin MSComDlg.CommonDialog CD_BancodeDados 
      Left            =   6435
      Top             =   1755
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.TextBox txtzip 
      Height          =   240
      Left            =   8100
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   1530
      Visible         =   0   'False
      Width           =   645
   End
   Begin VB.CommandButton Bt_BKP 
      Caption         =   "2 - Gerar Cópia de Segurança"
      Height          =   465
      Left            =   6210
      TabIndex        =   10
      Top             =   900
      Width           =   2490
   End
   Begin VB.Frame Frame1 
      Height          =   1545
      Left            =   45
      TabIndex        =   3
      Top             =   360
      Width           =   6045
      Begin VB.Label lblManutencao 
         Caption         =   "<Nenhum>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   14
         Top             =   1200
         Width           =   4515
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Manutenção:"
         Height          =   195
         Left            =   60
         TabIndex        =   13
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Local:"
         Height          =   195
         Left            =   945
         TabIndex        =   9
         Top             =   270
         Width           =   465
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tamanho:"
         Height          =   195
         Left            =   675
         TabIndex        =   8
         Top             =   585
         Width           =   735
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Criado/Modificado:"
         Height          =   195
         Left            =   45
         TabIndex        =   7
         Top             =   900
         Width           =   1365
      End
      Begin VB.Label Lb_Local 
         Caption         =   "<Nenhum>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   6
         Top             =   270
         Width           =   4545
      End
      Begin VB.Label Lb_Tamanho 
         Caption         =   "<Nenhum>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   5
         Top             =   585
         Width           =   4470
      End
      Begin VB.Label Lb_DtHr 
         Caption         =   "<Nenhum>"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   1440
         TabIndex        =   4
         Top             =   900
         Width           =   4515
      End
   End
   Begin VB.CommandButton Bt_CompactarBD 
      Caption         =   "1 - Reparar e Compactar BD"
      Height          =   465
      Left            =   6210
      TabIndex        =   2
      Top             =   405
      Width           =   2490
   End
   Begin VB.CommandButton Bt_Ok 
      Cancel          =   -1  'True
      Caption         =   "&Fechar"
      Height          =   465
      Left            =   6210
      TabIndex        =   0
      Top             =   1440
      Width           =   2490
   End
   Begin VB.Label Lb_Status 
      Height          =   195
      Left            =   90
      TabIndex        =   12
      Top             =   1980
      Width           =   6000
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "MANUTENÇÃO DO BANCO DE DADOS"
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
      TabIndex        =   1
      Top             =   0
      Width           =   8790
   End
End
Attribute VB_Name = "Form_InfoBD"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Tamanho As String
Private Sub HDBt(op As Boolean)
    Bt_CompactarBD.Enabled = op
    Bt_BKP.Enabled = op
    Bt_OK.Enabled = op
    MDIForm_Main.Enabled = op
End Sub


Private Sub Bt_BKP_Click()
    
    MsgBox "FUNÇÃO NÃO LIBERADA DEVIDO INSTABILIDADE.", vbInformation, "CESNet - Aviso"
    Exit Sub
    
    
    If MsgBox("Antes de gerar cópia de segurança é aconselhável Repara a Base de Dados." & vbCrLf & "Deseja continuar a gerar cópia de segurança?", vbInformation + vbYesNo, "CESNet") = vbNo Then
        Exit Sub
    End If
    
    
    Dim nmArq   As String
    Dim locArq  As String
    Dim caminho As String
    
    'CD_BancodeDados.DialogTitle = "Local e Nome do Arquivo?"
    'CD_BancodeDados.InitDir = App.Path
    'CD_BancodeDados.filename = "DB" & Format(Date, "yyyymmdd") & Format(Time, "hhmm")
    'CD_BancodeDados.Filter = "Zipado |*.zip"
    
    
    'CD_BancodeDados.ShowSave
    
    locArq = PathBD & "\Database\Dados\Dados.mdb" 'App.Path & "\database\dados\ilha\*.*"
    
    caminho = App.path & "\DB" & Format(Date, "yyyymmdd") & Format(Time, "hhmm")
    'caminho = Trim(CD_BancodeDados.filename)
    If caminho = "" Then Exit Sub
    
    
    
    HDBt (False)
    
    'INICIA BIBLIOTECA ZIP
    InicializaZip Me, txtzip
    
    DoEvents
    nmArq = caminho ' & "DB" & Format(Date, "yyyymmdd") & Format(Time, "hhmm") & ".zip"
    
    Compacta nmArq, locArq
    
    'nmArq = "lg" & Format(Date, "ddmmyyyy") & ".zip"
    locArq = PathBD & "\Database\Log\*.*"
    Compacta nmArq, locArq
    
    MsgBox "Cópia de segurança realizada com sucesso!", vbInformation, "CESNet"
    
    HDBt (True)
    Lb_Status.Caption = ""
    Call RegLog(UsuarioID, "COPIA DE SEGURANCA DO SISTEMA")
End Sub

Private Sub Bt_CompactarBD_Click()
    'DBEngine.RepairDatabase (PathBD & "\Dados.mdb") 'Repara o BD
    Dim RsTMP As Recordset
    On Error GoTo TratErro
    
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    MsgBox "Caso este Aplicativo esteja rodando em rede, as outras maquinas devem sair do aplicativo pois pode causar erro", vbInformation, "Aviso"
    Me.MousePointer = 11
    HDBt (False)
    
    BD.Close
    CompactarBD
'    DBEngine.CompactDatabase PathBD & "\Database\Dados\Dados.mdb", PathBD & "\Database\Dados\DadosBKP.mdb", , , ";PWD=k3bw82" 'Compacta o Banco de Dados Reparado e Renomeia.
'    'É importante compactar o BD apos repara-lo devido ao aumento do tamanho do mdb
'    Kill (PathBD & "\Database\Dados\Dados.mdb") 'Apaga o BD antigo
'    Name PathBD & "\Database\Dados\DadosBKP.mdb" As PathBD & "\Database\Dados\Dados.mdb" 'Renomeia o BD reparado para o corrente no sistema
    AbrirBD_DAO
    'Set RsTMP = BD.OpenRecordset("SELECT * FROM Config")
    'RsTMP.MoveFirst
    'RsTMP.Edit
    'RsTMP.Fields("UltManutencao") = Date
    'RsTMP.Update
    'RsTMP.Close


    'MsgBox "Correção concluida. Banco de Dados Restaurado.", vbExclamation, "AVISO"

    CarregaDados

    Me.MousePointer = 0
    HDBt (True)
    'Call RegLog(UsuarioID, "Tela:InfoBD - REPARACAO DE SEGURANCA DO SISTEMA")
    Exit Sub
TratErro:
    RegLogErros Err.Number, Err.Description, Me.Caption, 0
    MsgBox Err.Description & vbCrLf & "Operação cancelada.", vbInformation, "Erro n. " & Err.Number
    MDIForm_Main.Enabled = True
    AbrirBD_DAO
    Me.MousePointer = 0
    'Case 3356 - Banco de dados aberto por outro usuario em modo exclusivo
End Sub

Private Sub Bt_OK_Click()
    Unload Me
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    CarregaDados
    
End Sub

Private Sub CarregaDados()
    Tamanho = FileLen(PathBD & "\Database\Dados\Dados.mdb")
    PathBD = LCase(PathBD)
    Lb_Local.ToolTipText = PathBD
    Lb_Local.Caption = IIf(Len(PathBD) > 20, left(PathBD, 9) & "..." & Right(PathBD, 20), PathBD)
    Lb_Tamanho.Caption = Val(Tamanho) / 1024 & " Kb [" & Format(Tamanho, "###,###") & " bytes]"
    Lb_DtHr.Caption = FileDateTime(PathBD & "\Database\Dados\Dados.mdb")
    lblManutencao.Caption = ultManutencao
    'Me.Caption = DBEngine.Version
End Sub

Private Sub txtzip_Change()
    Lb_Status.Caption = "Compactando... " & GetPercentComplete(txtzip) & "%"
    DoEvents
End Sub


