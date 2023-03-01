VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Declaracoes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Declarações"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7920
   Icon            =   "Form_Declaracoes.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   5280
   ScaleWidth      =   7920
   Begin VB.Frame FrameFixa 
      Height          =   2595
      Left            =   60
      TabIndex        =   11
      Top             =   360
      Width           =   7755
      Begin VB.ListBox LstDoc 
         Enabled         =   0   'False
         Height          =   1035
         Left            =   360
         TabIndex        =   15
         Top             =   480
         Width           =   7155
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Declarações Fixas"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   14
         Top             =   1800
         Value           =   -1  'True
         Width           =   2115
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Declarações"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   180
         Width           =   2115
      End
      Begin VB.ComboBox Cb_Decl 
         Height          =   315
         ItemData        =   "Form_Declaracoes.frx":030A
         Left            =   360
         List            =   "Form_Declaracoes.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2100
         Width           =   7200
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2160
      Left            =   60
      TabIndex        =   0
      Top             =   3000
      Width           =   7800
      Begin MSComCtl2.DTPicker DTP_Dt 
         Height          =   315
         Left            =   1140
         TabIndex        =   10
         Top             =   1140
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   556
         _Version        =   393216
         Format          =   61800449
         CurrentDate     =   38959
      End
      Begin VB.ComboBox Cb_Ensino 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1140
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   660
         Width           =   3615
      End
      Begin VB.CommandButton Bt_Cancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   900
         Left            =   5640
         Picture         =   "Form_Declaracoes.frx":030E
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1140
         Width           =   2025
      End
      Begin VB.CommandButton Bt_Imprimir 
         Caption         =   "&Imprimir"
         Enabled         =   0   'False
         Height          =   900
         Left            =   3600
         Picture         =   "Form_Declaracoes.frx":0618
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   1140
         Width           =   2025
      End
      Begin MSMask.MaskEdBox MebMatricula 
         Height          =   330
         Left            =   1140
         TabIndex        =   3
         Top             =   240
         Width           =   1320
         _ExtentX        =   2328
         _ExtentY        =   582
         _Version        =   393216
         Appearance      =   0
         MaxLength       =   11
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Mask            =   "##.###.####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Caption         =   "Data da Declaração:"
         Height          =   495
         Left            =   195
         TabIndex        =   9
         Top             =   1020
         Width           =   915
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Curso:"
         Height          =   195
         Left            =   495
         TabIndex        =   7
         Top             =   720
         Width           =   615
      End
      Begin VB.Label Lb_Nome 
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   2580
         TabIndex        =   4
         Top             =   225
         Width           =   5100
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Matricula:"
         Height          =   195
         Left            =   420
         TabIndex        =   2
         Top             =   315
         Width           =   690
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "DECLARAÇÃO"
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
      Width           =   7965
   End
End
Attribute VB_Name = "Form_Declaracoes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim RsMatricula             As Recordset
Dim RsMatriculaEnsino       As Recordset
Dim RsMatriculaDisciplina   As Recordset
Dim RsMatriculaProva        As Recordset
Dim RsMatriculaHist         As Recordset
Dim RsTrafego               As Recordset
Dim RsEnsino                As Recordset

Dim Nome                    As String
Dim Endereco                As String
Dim Nasc                    As String
Dim EnsinoID                As Integer
Dim nArquivo                As String

Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub
Private Sub Bt_Imprimir_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    If EnsinoID = 0 Then
        MsgBox "Selecione um ensino!", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    
    If Option1.Item(0).Value = True Then
            If Trim(nArquivo) = "" Then
                MsgBox "Favor selecionar uma declaração!", vbInformation, "CESNet - Aviso"
                Exit Sub
            End If
            Call RelatorioWord(nArquivo, MebMatricula.Text, PgNomeEnsino(EnsinoID), DTP_Dt.Value)
            
        Else
    
            Select Case left(Cb_Decl.Text, 3)
                Case "001"
                    Call Rpt001(MebMatricula.Text, EnsinoID, DTP_Dt.Value)
                Case "002"
                    Call Rpt002(MebMatricula.Text, EnsinoID, DTP_Dt.Value)
                Case "003"
                    Call Rpt003(MebMatricula.Text, EnsinoID, DTP_Dt.Value)
                Case "004"
                    Call Rpt004(MebMatricula.Text, DTP_Dt.Value)
                Case "005"
                    Call Rpt005(MebMatricula.Text, EnsinoID, DTP_Dt.Value)
                Case Else
                    MsgBox "Declaração nao encontrada por favor avise ao suporte", vbInformation, "CESNet - Aviso"
                    Exit Sub
            End Select
    End If
End Sub




Private Sub Cb_Decl_DropDown()
    Cb_Decl.Clear
    Cb_Decl.AddItem ("001 - Declaração que esta cursando.")
    Cb_Decl.AddItem ("002 - Declaração de conclusão total.")
    Cb_Decl.AddItem ("003 - Declaração de conclusão parcial.")
    Cb_Decl.AddItem ("004 - Declaração de comparecimento.")
    Cb_Decl.AddItem ("005 - Declaração de provas efetuadas.")
End Sub

Private Sub Cb_Ensino_Click()
    EnsinoID = PgIDEnsino(Trim(Cb_Ensino.Text))
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.left = 0
    Me.top = 0
    Cb_Decl.Clear
    Cb_Ensino.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.EOF Then
            MsgBox "Erro ao carregar tabela de ensino.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_Ensino.AddItem (Trim(RsEnsino.Fields("Descr")))
                RsEnsino.MoveNext
            Loop
    End If
    
    DTP_Dt.Value = Date
    ListDocs
    
    
End Sub

Private Sub lstDoc_Click()
    nArquivo = PgNomeDoc(lstDoc.Text)
    
End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub
Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        '***** Checar Aviso ******
        If PgAviso(MebMatricula.Text) = True Then
            Exit Sub
        End If
        '*************************
        MebMatricula.PromptInclude = False
        If Trim(MebMatricula.Text) = "" Then
                MebMatricula.PromptInclude = True
                Exit Sub
            Else
                MebMatricula.PromptInclude = True
                MatrID = Trim(MebMatricula.Text)
        End If
        Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID = '" & MatrID & "'")
        If RsMatricula.BOF And RsMatricula.EOF Then
                MsgBox "Matricula não encontrada!", vbInformation, "CESNet - Aviso!"
                Lb_Nome.Caption = ""
                Cb_Decl.Enabled = False
                Cb_Ensino.Enabled = False
                Bt_Imprimir.Enabled = False
                MebMatricula.SetFocus
                Exit Sub
            Else
                RsMatricula.MoveFirst
                Nome = RsMatricula.Fields("Nome")
                Endereco = IIf(IsNull(RsMatricula.Fields("End")), "", RsMatricula.Fields("End"))
                Nasc = IIf(IsNull(RsMatricula.Fields("Nasc")), "", RsMatricula.Fields("Nasc"))
                Lb_Nome.Caption = Nome
                Cb_Decl.Enabled = True
                Cb_Ensino.Enabled = True
                Bt_Imprimir.Enabled = True
        End If
        RsMatricula.Close
    End If
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

Private Sub Option1_Click(Index As Integer)
    Select Case Index
        Case 0 'Declaracoes
        lstDoc.Enabled = True
        Cb_Decl.Enabled = False
        Case 1 'Declaracoes fixas
         lstDoc.Enabled = False
        Cb_Decl.Enabled = True
    End Select
End Sub
