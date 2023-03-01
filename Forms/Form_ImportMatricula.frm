VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Begin VB.Form Form_ImportMatricula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Importação de Matricula"
   ClientHeight    =   5265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7560
   Icon            =   "Form_ImportMatricula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7560
   Begin VB.Frame Frame4 
      Caption         =   "Status:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   5400
      TabIndex        =   18
      Top             =   2100
      Width           =   2055
      Begin VB.Label Lb_Status 
         Caption         =   "0"
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
         Left            =   900
         TabIndex        =   20
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Registros:"
         Height          =   195
         Left            =   120
         TabIndex        =   19
         Top             =   300
         Width           =   735
      End
      Begin VB.Image Image1 
         Height          =   480
         Left            =   1500
         Picture         =   "Form_ImportMatricula.frx":030A
         Top             =   480
         Width           =   480
      End
   End
   Begin MSComctlLib.ImageList IL_Import 
      Left            =   4920
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":0614
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":0930
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":0C4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":0F68
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":1284
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":15A0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":18BC
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_ImportMatricula.frx":1BD8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton Bt_ImprLogErro 
      Caption         =   "Imprimir Log Erro(s)"
      Height          =   495
      Left            =   5400
      TabIndex        =   15
      Top             =   1260
      Width           =   2055
   End
   Begin VB.Frame Frame3 
      Caption         =   "Log de Erro(s):"
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
      Left            =   60
      TabIndex        =   13
      Top             =   3180
      Width           =   7395
      Begin VB.ListBox Lst_LogErro 
         Height          =   1620
         Left            =   120
         TabIndex        =   14
         Top             =   240
         Width           =   7155
      End
   End
   Begin VB.CommandButton Bt_ImportarArq 
      Caption         =   "Importar"
      Height          =   495
      Left            =   5400
      TabIndex        =   12
      Top             =   600
      Width           =   2055
   End
   Begin VB.Frame Frame2 
      Caption         =   "Unidade Ensino:"
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
      TabIndex        =   6
      Top             =   2280
      Width           =   5055
      Begin VB.ComboBox Cb_UnidadeEnsino 
         Height          =   315
         Left            =   180
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   300
         Width           =   4755
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Local do Arquivo:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1755
      Left            =   60
      TabIndex        =   0
      Top             =   420
      Width           =   5055
      Begin MSComDlg.CommonDialog Cd_Conexao 
         Left            =   3120
         Top             =   300
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton Bt_LocArq 
         Height          =   435
         Left            =   4500
         Picture         =   "Form_ImportMatricula.frx":1EF4
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   180
         Width           =   495
      End
      Begin VB.TextBox Txt_LocArq 
         Height          =   285
         Left            =   180
         TabIndex        =   1
         Top             =   240
         Width           =   4275
      End
      Begin VB.Label Lb_DataArq 
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
         Height          =   255
         Left            =   1560
         TabIndex        =   17
         Top             =   1440
         Width           =   3255
      End
      Begin VB.Label Lb_Unidade 
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
         Left            =   1560
         TabIndex        =   16
         Top             =   1200
         Width           =   3315
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Data do Arquivo:"
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1440
         Width           =   1215
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Unidade Ensino:"
         Height          =   195
         Left            =   180
         TabIndex        =   10
         Top             =   1200
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome do Arquivo:"
         Height          =   195
         Left            =   120
         TabIndex        =   9
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Tamanho:"
         Height          =   195
         Left            =   600
         TabIndex        =   8
         Top             =   960
         Width           =   855
      End
      Begin VB.Label Lb_TamArq 
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
         Left            =   1560
         TabIndex        =   5
         Top             =   960
         Width           =   3315
      End
      Begin VB.Label Lb_NomeArq 
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
         Left            =   1560
         TabIndex        =   4
         Top             =   720
         Width           =   3315
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "IMPORTAÇÃO DE MATRICULA(S)"
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
      TabIndex        =   3
      Top             =   0
      Width           =   7575
   End
End
Attribute VB_Name = "Form_ImportMatricula"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUnidEnsino    As Recordset
Dim RsMatricula     As Recordset
Dim RsMatrEnsino    As Recordset
Dim RsMatrDiscipl   As Recordset

Dim Arq             As String
Dim caminho         As String
Dim TamArqKb        As Integer
Dim TamArqBy        As String

Dim CabData         As String
Dim CabUnidade      As String

Dim Dados(30)       As String
Dim UnidEns         As String
Dim UnidEnsNome     As String

Dim cStatus As Integer
Dim cReg As Integer

Private Sub PgDados(CamArq As String)
On Error GoTo TratErro
    Dim Arquivo As String
    Dim Linha As String
    Dim cont As Integer
    Dim TpReg As String
    Arquivo = FreeFile
    Open CamArq For Input As Arquivo
    cont = 0
    Do
        Line Input #Arquivo, Linha
        If Not cont = 0 Then
            If Trim(Linha) = "" Then Exit Do
            Linha = UCase(Trim(Linha))
            TpReg = Mid(Linha, 1, 3)
            DoEvents
            Call Status
            Select Case TpReg
                Case "MTR" 'Matricula
                    'Lb_Status.Caption = "Pegando dados de MATRICULA..."
                    Dados(0) = Val(Trim(Mid(Linha, 4, 15)))
                    
                    Dados(1) = Trim(Mid(Linha, 19, 8))
                    Dados(2) = Trim(Mid(Linha, 27, 50))
                    Dados(3) = Trim(Mid(Linha, 77, 50))
                    Dados(4) = Trim(Mid(Linha, 127, 30))
                    Dados(5) = Trim(Mid(Linha, 157, 30))
                    Dados(6) = Trim(Mid(Linha, 187, 2))
                    If PgNomeUF(Dados(6)) = "Sigla não cadastrada no CESNet" Then
                        Dados(6) = ""
                    End If
                    Dados(7) = Trim(Mid(Linha, 189, 9))
                    Dados(8) = Trim(Mid(Linha, 198, 1))
                    If Dados(8) <> "M" And Dados(8) <> "F" Then
                        Dados(8) = ""
                    End If
                    Dados(9) = Trim(Mid(Linha, 199, 8))
                    Dados(10) = Trim(Mid(Linha, 207, 30))
                    Dados(11) = Trim(Mid(Linha, 237, 14))
                    Dados(12) = Trim(Mid(Linha, 251, 14))
                    Dados(13) = Trim(Mid(Linha, 265, 14))
                    Dados(14) = Trim(Mid(Linha, 279, 20))
                    Dados(15) = Trim(Mid(Linha, 299, 20))
                    Dados(16) = Trim(Mid(Linha, 319, 20))
                    Dados(17) = Trim(Mid(Linha, 339, 30))
                    Dados(18) = EstCivil(Trim(Mid(Linha, 369, 2)))
                    Dados(19) = Trim(Mid(Linha, 371, 30))
                    Dados(20) = Trim(Mid(Linha, 401, 50))
                    Dados(21) = Trim(Mid(Linha, 451, 49))
                    If Chk_Matricula(Dados(0)) = "" Then
                            Call GrvMatr
                        Else
                            Call InfoLog(1, "Matricula n.: " & Dados(0) & " já cadastrada sob o Num.: " & Chk_Matricula(Dados(0)))
                            
                    End If
                Case "DCP" 'Diciplina
                    'Lb_Status.Caption = "Pegando dados de DISCIPLINA..."
                    Dados(0) = Trim(Mid(Linha, 4, 15))
                    Dados(1) = Trim(Mid(Linha, 19, 2))
                    Dados(1) = Val(UCase(Dados(1)))
                    Dados(2) = Trim(Mid(Linha, 21, 2))
                    Dados(2) = Val(Dados(2))
                    Dados(3) = Trim(Mid(Linha, 23, 8))
                    Dados(3) = Format(Dados(3), "0#/##/####")
                    Dados(4) = Trim(Mid(Linha, 31, 50))
                    Dados(5) = Trim(Mid(Linha, 81, 30))
                    Dados(6) = Trim(Mid(Linha, 111, 2))
                    If Chk_Matricula(Val(Dados(0))) = "" Then
                            Call InfoLog(2, "Matricula n.: " & Dados(0) & " não cadastrada.")
                            
                        Else
                            If PgNomeEnsino(Val(Dados(1))) = 0 Then
                                    Call InfoLog(2, "Matricula n.: " & Dados(0) & ". Ensino não cadastrada.")
                                Else
                                    If PgNomeDisciplina(Val(Dados(2))) = 0 Then
                                            Call InfoLog(2, " Matricula n.: " & Dados(0) & ". Disciplina: " & Dados(2) & " não encontrado.")
                                        Else
                                            Call GrvEnsino(Chk_Matricula(Dados(0)))
                                            Call GrvDiscipl(Chk_Matricula(Dados(0)))
                                    End If
                            End If
                    End If

                Case Else
                    'MsgBox "Dados inregulares"
                    Call InfoLog(3, "Erro ao coletar os dados.")
                    cont = 1
            End Select
            'Lb_Status.Caption = "<Nenhum>"
        End If
        cont = 1
    Loop
    cStatus = 0
    Call Status
    Close #Arquivo
    
    
    Exit Sub
TratErro:
    Select Case Err.Number
        Case 62 'Fim do Arquivo
            cStatus = 0
            Call Status
            Close #Arquivo
            Exit Sub
        Case Else
            'MsgBox Err.Description, vbInformation, "CESNet - Erro n. " & Err.Number
            Call InfoLog(3, "Erro: " & Err.Number & " - " & Err.Description)
            Resume Next
    End Select
    
End Sub

Private Sub Bt_ImportarArq_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    Lst_LogErro.Clear
    If UnidEns = "" Then
        MsgBox "Favor checar a Unidade de Ensino ao qual o arquivo será importado.", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    HDForm (False)
    PgDados (caminho & "\" & Arq)
    HDForm (True)
    cStatus = 0
    cReg = 0
End Sub

Private Sub Bt_ImprLogErro_Click()
    Dim Tmp As Integer
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    If Form_Impressora.LoadFormCI(True, True, False, False, False, True, False, False, False, False) = False Then
        Exit Sub
    End If
    'ObjPreview.Font = CI.Fonte
    'ObjPreview.CurrentY = 2000
    ObjPreview.FontSize = 8
    ObjPreview.CurrentY = 400
    ObjPreview.CurrentX = ObjPreview.ScaleWidth - (ObjPreview.TextWidth("CESNet [v." & Versao & "]") + 400)
    ObjPreview.Print "CESNet [v." & Versao & "]"
    ObjPreview.CurrentX = ObjPreview.ScaleWidth - (ObjPreview.TextWidth("CESNet [v." & Versao & "]") + 400)
    ObjPreview.Print "Data / Hora: " & Now
    ObjPreview.FontSize = 14
    ObjPreview.FontBold = True
    ObjPreview.CurrentY = 300
    ObjPreview.Print Tab(5); "Listagem dos Erros na Importação dos Dados"
    ObjPreview.FontItalic = True
    ObjPreview.FontSize = 10
    ObjPreview.Print
    ObjPreview.Print Tab(10); "Unidade de Ensino: "; Tab(35); Lb_Unidade.Caption
    ObjPreview.Print Tab(10); "Data do Arquivo: "; Tab(35); Lb_DataArq.Caption
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.Print
    ObjPreview.FontSize = 8
    For Tmp = 1 To Lst_LogErro.ListCount - 1
        ObjPreview.Print Tab(15); Lst_LogErro.List(Tmp)
    Next
    
    'Call ImpRodape
    
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
End Sub

Private Sub Bt_LocArq_Click()
    With CD_Conexao
        .DialogTitle = "CESNet - Conexão ao Banco de Dados"
        .InitDir = App.Path
        .Filter = "Texto|*.txt"
        '.FileName = "Dados.mdb"
        .DefaultExt = "*.txt"
        .Flags = cdlOFNFileMustExist Or cdlOFNHideReadOnly
        .ShowOpen
        If Len(Trim(.filename)) = 0 Then
            Exit Sub
        End If
        .filename = Left(.filename, Len(.filename) - Len(.FileTitle) - 1)
        If Len(.filename) >= 200 Then
                MsgBox "Nome do local para Arquivo muito extenso. Por favor modifique.", vbInformation, "CESNet - Aviso!"
                Txt_LocArq.SetFocus
                Exit Sub
            Else
                Arq = .FileTitle
                caminho = .filename
                
                TamArqBy = FileLen(caminho & "\" & Arq)
                TamArqKb = Val(TamArqBy) / 1024
                TamArqKb = IIf(TamArqKb = 0, 1, TamArqKb)
                Txt_LocArq.Text = caminho & "\" & Arq
                Lb_NomeArq.Caption = Arq
                Lb_TamArq.Caption = TamArqKb & " Kb [" & Format(TamArqBy, "###,###") & " bytes]"
                PgCabArq (caminho & "\" & Arq)
                Lb_Unidade.Caption = CabUnidade
                Lb_DataArq.Caption = CabData
        End If
    End With

End Sub

'Lb_Status = Matriculando aluno, atualizando matricula xxx
'***********************************
Private Sub PgCabArq(CamArq As String)
On Error GoTo TratErro
    Dim Arquivo As String
    Dim Linha As String
    Arquivo = FreeFile
    Open CamArq For Input As Arquivo
    Line Input #Arquivo, Linha
    'Linha = Crypto(PathBD)
    CabData = Mid(Linha, 1, 8)
    CabUnidade = Mid(Linha, 9, 50)
    Close #Arquivo
    If Not IsNumeric(CabData) Then
            MsgBox "Erro na checagem do cabeçalho", vbInformation, "CESNet - Erro"
            Call LimpDadosArq
            Call InfoLog(3, "Erro ao autenticar o cabeçalho do arquivo")
            Exit Sub
        Else
            CabData = Format(CabData, "0#/##/####")
            If Not IsDate(CabData) Then
                MsgBox "Data do cabeçalho invalida.", vbInformation, "CESNet - Aviso"
                Call InfoLog(3, "Erro ao autenticar o cabeçalho do arquivo")
                Call LimpDadosArq
                Exit Sub
            End If
    End If
    Exit Sub
TratErro:
    MsgBox Err.Description, vbInformation, "CESNet - Erro n. " & Err.Number
    Resume Next
End Sub

Private Sub Cb_UnidadeEnsino_Click()
    UnidEns = Trim(Left(Cb_UnidadeEnsino.Text, 3))
    UnidEnsNome = Trim(Cb_UnidadeEnsino.Text)
End Sub

Private Sub Cb_UnidadeEnsino_DropDown()
    Cb_UnidadeEnsino.Clear
    Set RsUnidEnsino = BD.OpenRecordset("SELECT * FROM Unidades ORDER BY UnidID")
    If RsUnidEnsino.BOF And RsUnidEnsino.EOF Then
            Exit Sub
        Else
            With RsUnidEnsino
                .MoveFirst
                Do Until .EOF
                    Cb_UnidadeEnsino.AddItem (.Fields("UnidID") & " - " & .Fields("Nome"))
                    .MoveNext
                Loop
            End With
            
    End If
    
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.Top = 0
    Me.Left = 0
    'Call Status
End Sub
Private Sub Status()
    cReg = cReg + 1
    Lb_Status.Caption = cReg
    cStatus = cStatus + 1
    cStatus = IIf(cStatus = 9, 1, cStatus)
    Image1.Picture = IL_Import.ListImages.Item(cStatus).Picture
End Sub
Private Sub Txt_LocArq_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub HDForm(op As Boolean)
    Txt_LocArq.Enabled = op
    Bt_LocArq.Enabled = op
    Cb_UnidadeEnsino.Enabled = op
    
    Bt_ImportarArq.Enabled = op
    Bt_ImprLogErro.Enabled = op
'    Bt_Cancelar.Enabled = Op
End Sub
Private Sub LimpDadosArq()
    
    CabData = "<Nenhum>"
    CabUnidade = "<Nenhum>"
    Arq = ""
    caminho = ""
    TamArqKb = 0
    TamArqBy = ""
    Lb_NomeArq.Caption = "<Nenhum>"
    Lb_TamArq.Caption = "<Nenhum>"
    Lb_Unidade.Caption = "<Nenhum>"
    Lb_DataArq.Caption = "<Nenhum>"
    Txt_LocArq.Text = ""
End Sub
Private Function EstCivil(EstID As String) As String
    Select Case EstID
        Case "01"
            EstCivil = "AMIGADO(A)"
        Case "02"
            EstCivil = "CASADO(A)"
        Case "03"
            EstCivil = "DISQUITADO(A)"
        Case "04"
            EstCivil = "SOLTEIRO(A)"
        Case "05"
            EstCivil = "VIUVO(A)"
        Case "06"
            EstCivil = "OUTROS"
        Case Else
            EstCivil = ""
    End Select
End Function
Private Function Chk_Matricula(NumAnt As String) As String
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE Unidade LIKE '" & UnidEns & "*' AND NumAnt = '" & Val(NumAnt) & "'")
    If RsMatricula.BOF And RsMatricula.EOF Then
            Chk_Matricula = ""
        Else
            RsMatricula.MoveFirst
            Chk_Matricula = RsMatricula.Fields("MatrID")
    End If
End Function
Private Sub GrvMatr()
    Dim MatrID As String
    'Checa se existem alunos com o mesmo nome
    Dim RsTmp As Recordset
    
    Do Until InStr(Dados(2), "'") = 0
        Dados(2) = Left(Dados(2), InStr(Dados(2), "'") - 1) & " " & Right(Dados(2), Len(Dados(2)) - InStr(Dados(2), "'"))
    Loop
    
    Set RsTmp = BD.OpenRecordset("SELECT Matriculas.Nome FROM Matriculas WHERE Nome = '" & Dados(2) & "'")
    If RsTmp.BOF And RsTmp.EOF Then
            RsTmp.Close
        Else
            Dados(2) = Dados(2) & "."
            Call InfoLog(1, "Matricula: " & Dados(0) & " - " & Dados(2) & " - Nome Duplicado. Foi incluso ponto<.> no final do nome.")
            RsTmp.Close
    End If
    
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE Unidade LIKE '" & UnidEns & "*' ORDER BY MatrID")
    If RsMatricula.BOF And RsMatricula.EOF Then
            MatrID = Right(Date, 2) & "." & UnidEns & ".0001"
        Else
            With RsMatricula
                RsMatricula.MoveLast
                MatrID = Right(RsMatricula.Fields("MatrID"), 4)
                MatrID = Val(MatrID) + 1
                
                MatrID = Right(Date, 2) & "." & UnidEns & "." & Left(String(4, "0"), 4 - Len(Trim(MatrID))) & Trim(MatrID)
            End With
    End If
    With RsMatricula
        .AddNew
        .Fields("MatrID") = MatrID
        .Fields("Nome") = Dados(2)
        .Fields("Unidade") = UnidEnsNome
        .Fields("DtMat") = IIf(Trim(Dados(1)) = "", Date, Format(Dados(1), "0#/##/####"))
        .Fields("End") = IIf(Trim(Dados(3)) = "", Null, Trim(Dados(3)))
        .Fields("Bai") = IIf(Trim(Dados(4)) = "", Null, Trim(Dados(4)))
        .Fields("Mun") = IIf(Trim(Dados(5)) = "", Null, Trim(Dados(5)))
        .Fields("UF") = IIf(Trim(Dados(6)) = "", Null, Trim(Dados(6)))
        .Fields("CEP") = IIf(Trim(Dados(7)) = "", Null, Trim(Dados(7)))
        .Fields("Sexo") = IIf(Trim(Dados(8)) = "", Null, Trim(Dados(8)))
                
        .Fields("Nasc") = IIf(Trim(Dados(9)) = "", Null, Format(Dados(9), "0#/##/####"))
                
        .Fields("Mail") = IIf(Trim(Dados(10)) = "", Null, Trim(Dados(10)))
        .Fields("Cel") = IIf(Trim(Dados(11)) = "", Null, Trim(Dados(11)))
        .Fields("Tel1") = IIf(Trim(Dados(12)) = "", Null, Trim(Dados(12)))
        .Fields("Tel2") = IIf(Trim(Dados(13)) = "", Null, Trim(Dados(13)))
    
        .Fields("RG") = IIf(Trim(Dados(14)) = "", Null, Trim(Dados(14)))
        .Fields("OE") = IIf(Trim(Dados(15)) = "", Null, Trim(Dados(15)))
        .Fields("CertNasc") = IIf(Trim(Dados(16)) = "", Null, Trim(Dados(16)))
        .Fields("Natural") = IIf(Trim(Dados(17)) = "", Null, Trim(Dados(17)))
        .Fields("EstCivil") = IIf(Trim(Dados(18)) = "", Null, Trim(Dados(18)))
        .Fields("Nacion") = IIf(Trim(Dados(19)) = "", Null, Trim(Dados(19)))
                    
        .Fields("NumAnt") = IIf(Trim(Dados(0)) = "", Null, Dados(0))
                    
        .Fields("Mae") = IIf(Trim(Dados(20)) = "", Null, Trim(Dados(20)))
        .Fields("Pai") = IIf(Trim(Dados(21)) = "", Null, Trim(Dados(21)))
                
        .Fields("UsuarioID") = UsuarioID
        .Fields("DtHrSis") = Now
        .Update
        '.FindFirst "MatrID = '" & MatrID & "'"
        'Do Until .NoMatch = False
        '    .FindFirst "MatrID = '" & MatrID & "'"
        'Loop
        .Close
        
    End With
End Sub

Private Sub GrvDiscipl(MatrID As String)
    Set RsMatrDiscipl = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & Dados(1) & " AND DisciplinaID = " & Dados(2))
    If RsMatrDiscipl.BOF And RsMatrDiscipl.EOF Then
            With RsMatrDiscipl
                .AddNew
                .Fields("MatrID") = MatrID
                .Fields("EnsinoID") = Dados(1)
                .Fields("DisciplinaID") = Dados(2)
                .Fields("DtConclusao") = Dados(3)
                .Fields("Local") = Dados(4)
                .Fields("Cidade") = IIf(Trim(Dados(5)) = "", Null, Dados(5))
                .Fields("UF") = IIf(Trim(Dados(6)) = "", Null, Dados(6))
                .Update
                .FindFirst "MatrID = '" & MatrID & "'"
                Do Until .NoMatch = False
                    .FindFirst "MatrID = '" & MatrID & "'"
                Loop
            End With
            If Chk_ConcEnsino(MatrID, CInt(Dados(1))) = True Then
                Call ConcEnsino(MatrID, CInt(Dados(1)))
            End If
        Else
            Call InfoLog(2, "Matricula: " & MatrID & " / Disciplina: " & PgNomeDisciplina(Val(Dados(2))) & " já cadastrada.")
    End If
End Sub
Private Sub InfoLog(Tipo As Integer, msg As String)
    Select Case Tipo
        Case 1 'Matricula
            Lst_LogErro.AddItem "[Matricula] - " & msg
        Case 2 'Disciplina
            Lst_LogErro.AddItem "[Disciplina] - " & msg
        Case 3 'Geral
            Lst_LogErro.AddItem "[Geral] - " & msg
    End Select
End Sub
Private Sub GrvEnsino(MatrID As String)
    Set RsMatrEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & Dados(1))
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
                RsMatrEnsino.AddNew
                RsMatrEnsino.Fields("MatrID") = MatrID
                RsMatrEnsino.Fields("EnsinoID") = Dados(1)
                RsMatrEnsino.Fields("DtInicio") = Dados(3)
                RsMatrEnsino.Update
                RsMatrEnsino.FindFirst "MatrID = '" & MatrID & "'"
                Do Until RsMatrEnsino.NoMatch = False
                    RsMatrEnsino.FindFirst "MatrID = '" & MatrID & "'"
                Loop
        Else
                RsMatrEnsino.MoveFirst
                If RsMatrEnsino.Fields("DtInicio") > Dados(3) Then
                    RsMatrEnsino.Edit
                    RsMatrEnsino.Fields("DtInicio") = Dados(3)
                    RsMatrEnsino.Fields("DtFinal") = Null
                    RsMatrEnsino.Update
                End If
    End If
End Sub


Private Sub ConcEnsino(MatrID As String, EnsinoID As Integer)
    Set RsMatrDiscipl = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " ORDER BY DtConclusao ASC")
    If RsMatrDiscipl.BOF And RsMatrDiscipl.EOF Then
            MsgBox "Erro ao localizar a data de conclusao do Ensino na Tabela Disciplina", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsMatrDiscipl.MoveLast
    End If
    Set RsMatrEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
            MsgBox "Erro ao localizar o ensino na Tabela Disciplina.", vbInformation, "CESNet - Aviso"
        Else
            RsMatrEnsino.MoveFirst
            RsMatrEnsino.Edit
            RsMatrEnsino.Fields("DtFinal") = RsMatrDiscipl.Fields("DtConclusao")
            RsMatrEnsino.Fields("Local") = RsMatrDiscipl.Fields("Local")
            RsMatrEnsino.Update
    End If
End Sub
