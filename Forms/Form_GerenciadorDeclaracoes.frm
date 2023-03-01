VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form_GerenciadorDeclaracoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Gerenciador de Declarações"
   ClientHeight    =   5910
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9270
   Icon            =   "Form_GerenciadorDeclaracoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5910
   ScaleWidth      =   9270
   Begin MSComctlLib.Toolbar tbMenu 
      Align           =   1  'Align Top
      Height          =   870
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   1535
      ButtonWidth     =   3069
      ButtonHeight    =   1376
      Appearance      =   1
      ImageList       =   "ilImagem"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Capturar Documento"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Armazenar Documento"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Style           =   3
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Caption         =   "Excluir Documento"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList ilImagem 
         Left            =   6060
         Top             =   120
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   32
         ImageHeight     =   32
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   3
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_GerenciadorDeclaracoes.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_GerenciadorDeclaracoes.frx":0FE4
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_GerenciadorDeclaracoes.frx":12FE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Documentos:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2235
      Left            =   120
      TabIndex        =   8
      Top             =   1260
      Width           =   9015
      Begin VB.ListBox lstDoc 
         Height          =   1815
         Left            =   180
         TabIndex        =   10
         Top             =   240
         Width           =   8655
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Capturar:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2295
      Left            =   120
      TabIndex        =   1
      Top             =   3540
      Width           =   9015
      Begin VB.TextBox txtDescr 
         Enabled         =   0   'False
         Height          =   315
         Left            =   180
         MaxLength       =   50
         TabIndex        =   6
         Top             =   1800
         Width           =   8655
      End
      Begin MSComDlg.CommonDialog cdGD 
         Left            =   7680
         Top             =   1320
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "CESNet - Capturar Documento"
         FileName        =   "*.doc"
      End
      Begin VB.Label Label1 
         Caption         =   "Título do Documento:"
         Height          =   195
         Left            =   180
         TabIndex        =   7
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label lbNomeArq 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   180
         TabIndex        =   5
         Top             =   600
         Width           =   8715
      End
      Begin VB.Label lbOrigem 
         BorderStyle     =   1  'Fixed Single
         Height          =   315
         Left            =   180
         TabIndex        =   4
         Top             =   1200
         Width           =   8715
      End
      Begin VB.Label Label3 
         Caption         =   "Local de Origem:"
         Height          =   255
         Left            =   180
         TabIndex        =   3
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label2 
         Caption         =   "Nome do Arquivo:"
         Height          =   255
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   1875
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "GERENCIADOR DE DECLARAÇÕES"
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
      Top             =   900
      Width           =   9315
   End
End
Attribute VB_Name = "Form_GerenciadorDeclaracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim Origem      As String
Dim Destino     As String
Dim nArquivo    As String

Private Sub GrvDados()
    On Error GoTo TrtErro
    Dim RsTMP As Recordset

    If Trim(Origem) = "" Or Trim(Destino) = "" Then Exit Sub
    If txtDescr.Text = "" Then
        MsgBox "Favor colocar o título do Documento", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    FileCopy Origem, Destino
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Documentos")
    RsTMP.AddNew
    RsTMP.Fields("nArq") = nArquivo
    RsTMP.Fields("Descr") = Trim(txtDescr.Text)
    RsTMP.Update
    
    lbOrigem.Caption = ""
    txtDescr.Text = ""
    lbNomeArq.Caption = ""
    txtDescr.Enabled = False
    tbMenu.Buttons.Item(2).Enabled = False
    tbMenu.Buttons.Item(4).Enabled = False
    Exit Sub
TrtErro:
    MsgBox Err.Description, vbInformation, Err.Number
    If Dir(Destino) <> "" Then
        Kill Destino
    End If
End Sub

Private Sub CapDoc()
    On Error GoTo TrtErro
    
    cdGD.FileName = ""
    cdGD.DialogTitle = "CESNet - Gerenciador de Declarações"
    cdGD.InitDir = App.path
    cdGD.filter = "Documento do Word|*.doc"
    cdGD.DefaultExt = "*.doc"
    cdGD.Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
    cdGD.ShowOpen
    Origem = cdGD.FileName
    If Trim(Origem) = "" Then
        lbOrigem.Caption = ""
        txtDescr.Text = ""
        lbNomeArq.Caption = ""
        txtDescr.Enabled = False
        tbMenu.Buttons.Item(2).Enabled = False
        tbMenu.Buttons.Item(4).Enabled = False
        Exit Sub
    End If
    
    nArquivo = Format(Date, "YYYYMMDD") & Format(Time, "HHMMSS") & ".doc"
    Destino = PathBD & "\Database\RptModelo\" & nArquivo
    
    lbNomeArq.Caption = nArquivo
    lbOrigem.Caption = Origem
    txtDescr.Text = ""
    txtDescr.Enabled = True
    tbMenu.Buttons.Item(2).Enabled = True
    tbMenu.Buttons.Item(4).Enabled = False
    Exit Sub
TrtErro:
    MsgBox Err.Description, vbInformation, Err.Number
End Sub

Private Sub ListDocs()
    Dim RsTMP As Recordset
    lstDoc.Clear
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Documentos ORDER BY Descr")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                lstDoc.AddItem RsTMP.Fields("Descr")
                RsTMP.MoveNext
            Loop
            RsTMP.Close
    End If
End Sub

Private Sub Form_Load()
    ListDocs
End Sub
Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub lstDoc_Click()
    Dim DocDescr As String
    
    DocDescr = lstDoc.Text
    nArquivo = PgNomeDoc(DocDescr)
    lbNomeArq.Caption = nArquivo
    lbOrigem.Caption = ""
    txtDescr.Text = DocDescr
    tbMenu.Buttons.Item(4).Enabled = True
End Sub

Private Sub tbMenu_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Index
        Case 1 'Capturar Documento
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            CapDoc
        Case 2 'Armazenar Doc
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            If ValidarSoftware("Documentos") = False Then Exit Sub
            GrvDados
        Case 4 'Excluir
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            ApagarArq
            tbMenu.Buttons.Item(4).Enabled = False
    End Select
    ListDocs
End Sub
Private Sub ApagarArq()
    On Error Resume Next
    'Altenticar usuario
    If Form_AutenticacaoUsuario.CarregarForm = False Then
        Exit Sub
    End If

    BD.Execute "DELETE * FROM Documentos WHERE nArq = '" & nArquivo & "'"
    Kill PathBD & "\Database\RptModelo\" & nArquivo
    
    

    MsgBox "Arquivo Excluido!", vbInformation, "CESNet - Aviso"
    Call RegLog("DECLARACOES", "EXCLUIU DECLARAÇÃO: " & txtDescr.Text)
    lbOrigem.Caption = ""
    txtDescr.Text = ""
    lbNomeArq.Caption = ""
End Sub
