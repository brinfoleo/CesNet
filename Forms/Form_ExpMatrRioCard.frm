VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_ExpMatrRioCard 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Exportação RioCard"
   ClientHeight    =   2205
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6240
   Icon            =   "Form_ExpMatrRioCard.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2205
   ScaleWidth      =   6240
   Begin VB.Frame Frame1 
      Height          =   1710
      Left            =   90
      TabIndex        =   1
      Top             =   360
      Width           =   6045
      Begin MSComDlg.CommonDialog CD_Riocard 
         Left            =   3120
         Top             =   0
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Bt_Cancelar 
         Cancel          =   -1  'True
         Caption         =   "Cancelar"
         Height          =   510
         Left            =   4170
         TabIndex        =   8
         Top             =   885
         Width           =   1755
      End
      Begin VB.Frame Frame3 
         Caption         =   "Status:"
         Height          =   1455
         Left            =   105
         TabIndex        =   3
         Top             =   105
         Width           =   3990
         Begin VB.Label Lb_RegGerados 
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2700
            TabIndex        =   7
            Top             =   810
            Width           =   690
         End
         Begin VB.Label Lb_RegSel 
            Caption         =   "0000"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2700
            TabIndex        =   6
            Top             =   270
            Width           =   735
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Registros Gravados:"
            Height          =   240
            Left            =   1140
            TabIndex        =   5
            Top             =   765
            Width           =   1455
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Total de Reg. na Base de Dados:"
            Height          =   240
            Left            =   30
            TabIndex        =   4
            Top             =   270
            Width           =   2565
         End
      End
      Begin VB.CommandButton Bt_GerarArquivo 
         Caption         =   "Gerar Arquivo"
         Height          =   510
         Left            =   4170
         TabIndex        =   2
         Top             =   255
         Width           =   1755
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "EXPORTAÇÃO RIOCARD"
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
      Width           =   6300
   End
End
Attribute VB_Name = "Form_ExpMatrRioCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatricula As Recordset

Private Sub GrvDados(caminho As String, txtDados As String)
On Error GoTo TrtErro
'Registra oq os usuarios do sistema estao fazendo

'define o ObjPreview filesystem e demais variaveis
Dim fso As New FileSystemObject
Dim Arquivo As File
Dim arquivoLog As TextStream


'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
    arquivoLog.WriteLine txtDados
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
    Resume Next
End Sub

Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub
Private Sub CarregarDados()
    Me.MousePointer = 11
    HD_Bt (False)
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas") ' WHERE DtMat >= #" & Format(DTP_DtInicial.Value, "mm/dd/yyyy") & "# AND DtMat <= #" & Format(DTP_DtFinal.Value, "mm/dd/yyyy") & "#")
    If RsMatricula.BOF And RsMatricula.EOF Then
            Lb_RegSel.Caption = "0000"
            Lb_RegGerados.Caption = "0000"
            HD_Bt (False)
            RsMatricula.Close
        Else
            RsMatricula.MoveLast
            Lb_RegSel.Caption = RsMatricula.RecordCount
            Lb_RegGerados.Caption = "0000"
            RsMatricula.MoveFirst
            HD_Bt (True)
    End If
    Me.MousePointer = 0
    
End Sub


Private Sub Bt_GerarArquivo_Click()
    Dim caminho As String
    Dim Dados   As String
    Dim cont As Integer
    CD_Riocard.DialogTitle = "Local e Nome do Arquivo?"
    CD_Riocard.InitDir = App.Path
    CD_Riocard.filename = "RC" & Format(Date, "yyyy") & PgDadosUnid(UnidadeEnsino).CodEscolar
    CD_Riocard.Filter = "Texto |*.txt"
    'CD_Riocard.Filter = "Todos | *.*"
   
    CD_Riocard.ShowSave
   
    caminho = Trim(CD_Riocard.filename)
    If Trim(caminho) = "" Then Exit Sub
    
    cont = 0
  
    Bt_GerarArquivo.Enabled = False
    Bt_Cancelar.Enabled = False
    Me.MousePointer = 11
    
    RsMatricula.MoveFirst
    Call GrvDados(caminho, "Escola;CODESCOLA;MatriculaAluno;SERIE;TURNO;Aluno;PAI;MAE;SEXO;DATANASC;TPLOGRARES;LOGRARES;NUMERORES;COMPRES;BAIRRORES;MUNICIPIORES;CEPRES;UFRES;NRO_REGUA")
    Do Until RsMatricula.EOF
        DoEvents
        
        Dados = PgDadosUnid(UnidadeEnsino).Nome & ";" & PgDadosUnid(UnidadeEnsino).CodEscolar & ";" & RsMatricula.Fields("MatrID") & _
                ";1;4" & ";" & RsMatricula.Fields("Nome") & ";" & _
                RsMatricula.Fields("Pai") & ";" & RsMatricula.Fields("Mae") & ";" & _
                RsMatricula.Fields("sexo") & ";" & RsMatricula.Fields("Nasc") & ";" & _
                ";" & RsMatricula.Fields("End") & ";" & _
                RsMatricula.Fields("Numero") & ";" & RsMatricula.Fields("Compl") & ";" & _
                RsMatricula.Fields("Bai") & ";" & RsMatricula.Fields("Mun") & ";" & _
                RsMatricula.Fields("CEP") & ";" & RsMatricula.Fields("UF") & ";"
        If PgStatusMatricula(RsMatricula.Fields("MatrID")) = "ATIVO" Then
            cont = cont + 1
            Lb_RegGerados.Caption = cont
            Call GrvDados(caminho, Dados)
        End If
        RsMatricula.MoveNext
    Loop

    Bt_GerarArquivo.Enabled = True
    Bt_Cancelar.Enabled = True
    Me.MousePointer = 0
End Sub




Private Sub Form_Load()
    
    CarregarDados
End Sub
Private Sub HD_Bt(op As Boolean)
    Bt_GerarArquivo.Enabled = op
End Sub
