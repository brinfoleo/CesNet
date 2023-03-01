VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Professores 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Professores"
   ClientHeight    =   6795
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7425
   Icon            =   "Form_Professores.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6795
   ScaleWidth      =   7425
   Begin VB.ComboBox Cb_Chv 
      Enabled         =   0   'False
      Height          =   315
      Left            =   5400
      Style           =   2  'Dropdown List
      TabIndex        =   19
      Top             =   1440
      Width           =   1725
   End
   Begin VB.ComboBox Cb_Nome 
      Height          =   315
      Left            =   720
      TabIndex        =   18
      Top             =   1035
      Width           =   6495
   End
   Begin VB.Frame Frame1 
      Caption         =   "Disciplinas:"
      Height          =   3315
      Left            =   90
      TabIndex        =   13
      Top             =   3420
      Width           =   7215
      Begin VB.ListBox Lst_Disciplinas 
         Enabled         =   0   'False
         Height          =   2985
         Left            =   120
         Sorted          =   -1  'True
         Style           =   1  'Checkbox
         TabIndex        =   14
         Top             =   240
         Width           =   6975
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Endereço:"
      Height          =   1515
      Left            =   60
      TabIndex        =   2
      Top             =   1800
      Width           =   7215
      Begin VB.ComboBox CB_UF 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form_Professores.frx":030A
         Left            =   6300
         List            =   "Form_Professores.frx":035F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   660
         Width           =   750
      End
      Begin VB.TextBox Txt_Mun 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3960
         MaxLength       =   30
         TabIndex        =   5
         Top             =   660
         Width           =   1815
      End
      Begin VB.TextBox Txt_Bai 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   4
         Top             =   660
         Width           =   1995
      End
      Begin VB.TextBox Txt_End 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   50
         TabIndex        =   3
         Top             =   240
         Width           =   6135
      End
      Begin MSMask.MaskEdBox Meb_Cep 
         Height          =   255
         Left            =   900
         TabIndex        =   7
         Top             =   1080
         Width           =   1035
         _ExtentX        =   1826
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   9
         Mask            =   "#####-###"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Meb_Nasc 
         Height          =   315
         Left            =   5820
         TabIndex        =   15
         Top             =   1080
         Width           =   1215
         _ExtentX        =   2143
         _ExtentY        =   556
         _Version        =   393216
         Enabled         =   0   'False
         MaxLength       =   10
         Format          =   "DD/MM/YYYY"
         Mask            =   "##/##/####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Nascimento:"
         Height          =   195
         Left            =   4740
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP:"
         Height          =   195
         Left            =   480
         TabIndex        =   12
         Top             =   1140
         Width           =   375
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   195
         Left            =   6000
         TabIndex        =   11
         Top             =   720
         Width           =   255
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "Municipio:"
         Height          =   195
         Left            =   3120
         TabIndex        =   10
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   420
         TabIndex        =   9
         Top             =   780
         Width           =   435
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   8
         Top             =   300
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   20
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1058
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "IL_Menu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alterar"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList IL_Menu 
         Left            =   6255
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   32
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":03CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":06E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":0A02
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":0D1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":1036
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":1350
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":166A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":1984
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":1C9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":1FB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":22D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":25EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":2906
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":2C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":2F3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":3254
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":356E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":3888
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":3BA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":3EBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":41D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":44F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":480A
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":4B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":4E3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":5158
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":5472
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":578C
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":5AA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":5DC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":60DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Professores.frx":63F4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Chave:"
      Height          =   195
      Left            =   4620
      TabIndex        =   17
      Top             =   1500
      Width           =   735
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   195
      Left            =   120
      TabIndex        =   1
      Top             =   1080
      Width           =   495
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "PROFESSORES"
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
      Top             =   600
      Width           =   7455
   End
End
Attribute VB_Name = "Form_Professores"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsProfessor As Recordset
Dim RsDisciplinas As Recordset
Dim RsProfessorDisciplina As Recordset
Dim RsUsuario As Recordset
Dim RsTMP As Recordset

Dim UsuID As Integer
Dim ProfessorID As String
Dim tmp As String
Dim Acao As Integer
Dim cont As Integer






Private Sub Cb_Chv_DropDown()
    Cb_Chv.Clear
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario ORDER BY Chv")
    If RsUsuario.BOF And RsUsuario.EOF Then
        Exit Sub
    End If
    RsUsuario.MoveFirst
    Do Until RsUsuario.EOF
        Cb_Chv.AddItem (RsUsuario.Fields("Chv"))
        RsUsuario.MoveNext
    Loop
End Sub

Private Sub Cb_Nome_DropDown()
    If Acao = 1 Or Acao = 2 Then Exit Sub
    Cb_Nome.Clear
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores ORDER BY Nome")
    If RsProfessor.BOF And RsProfessor.EOF Then
            Exit Sub
        Else
            RsProfessor.MoveFirst
            Cb_Nome.Clear
            Do Until RsProfessor.EOF
                Cb_Nome.AddItem (RsProfessor.Fields("Nome"))
                RsProfessor.MoveNext
            Loop
    End If
End Sub



Private Sub LimpForm()
    Cb_Chv.Clear
    Cb_Nome.Text = ""
    Txt_End.Text = ""
    Txt_Bai.Text = ""
    Txt_Mun.Text = ""
    CB_UF = " "
    Meb_Nasc.PromptInclude = False
    Meb_Nasc.Text = ""
    Meb_Nasc.PromptInclude = True
    Lst_Disciplinas.Clear
End Sub
Private Sub HDForm(op As Boolean)
    Cb_Chv.Enabled = op
    'Cb_Nome.Enabled = op
    Txt_End.Enabled = op
    Txt_Bai.Enabled = op
    Txt_Mun.Enabled = op
    CB_UF.Enabled = op
    MEB_Cep.Enabled = op
    Meb_Nasc.Enabled = op
    Lst_Disciplinas.Enabled = op
End Sub





Private Sub GrvProf()
  If Cb_Nome.Text = "" Then 'Or Cb_Chv.Text = ""
        MsgBox "Os campos: NOME não podem ser deixados em branco." & Chr(13) & "Por favor Verifique!", vbInformation, "CESNet - Aviso!"
        Cb_Nome.SetFocus
        Exit Sub
    End If
    'UsuID = Cb_Chv.Text
    'Set RsTmp = BD.OpenRecordset("SELECT * FROM Usuario WHERE Chv = '" & ProfessorID & "'")
    'If RsTmp.BOF And RsTmp.EOF Then
    '    MsgBox "Chave não cadastrada como USUÁRIO do sistema. Por favor verifique!", vbInformation ,"CESNet - Aviso!"
    '    Exit Sub
    'End If
    If Cb_Chv.Text = Usuario Then
        MsgBox "O Usuario não pode cadastrar-se como professor.", vbInformation, "CESNet - Atenção"
        Exit Sub
    End If
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores")
    With RsProfessor
        'ProfessorID = Cb_Chv.Text
        Select Case Acao
            Case 1
                UsuID = PgUsuIDID(IIf(Trim(Cb_Chv.Text) = "", 0, Cb_Chv.Text))
                If UsuID <> 0 Then
                    RsProfessor.FindFirst "UsuarioID = " & UsuID
                    If RsProfessor.NoMatch Then
                        Else
                            MsgBox "Usuario: " & Cb_Chv.Text & " já cadastrado como Professor. Por favor Verifique!", vbInformation, "CESNet -  Aviso!"
                            Cb_Chv.SetFocus
                            Exit Sub
                    End If
                End If
                'tmp = .Fields("MatrID")
                .AddNew
                '.Fields("ProfID") = ProfessorID
                .Fields("UsuarioID") = UsuID
                .Fields("Nome") = Cb_Nome.Text
                .Fields("End") = IIf(Txt_End.Text = "", " ", Txt_End.Text)
                .Fields("Bai") = IIf(Txt_Bai.Text = "", " ", Txt_Bai.Text)
                .Fields("Mun") = IIf(Txt_Mun.Text = "", " ", Txt_Mun.Text)
                .Fields("UF") = IIf(CB_UF.Text = "", " ", CB_UF.Text)
                .Fields("CEP") = IIf(MEB_Cep.Text = "", " ", MEB_Cep.Text)
                Meb_Nasc.PromptInclude = False
                If Trim(Meb_Nasc.Text) = "" Then
                    Else
                    Meb_Nasc.PromptInclude = True
                    .Fields("Nasc") = Meb_Nasc.Text
                End If
                Meb_Nasc.PromptInclude = True
                .Update
                .MoveLast
                ProfessorID = .Fields("ProfID")
            Case 2
                UsuID = PgUsuIDID(IIf(Trim(Cb_Chv.Text) = "", 0, Cb_Chv.Text))
                .FindFirst "ProfID = " & ProfessorID
                .Edit
                .Fields("UsuarioID") = UsuID
                .Fields("Nome") = Cb_Nome.Text
                .Fields("End") = IIf(Txt_End.Text = "", " ", Txt_End.Text)
                .Fields("Bai") = IIf(Txt_Bai.Text = "", " ", Txt_Bai.Text)
                .Fields("Mun") = IIf(Txt_Mun.Text = "", " ", Txt_Mun.Text)
                .Fields("UF") = IIf(CB_UF.Text = "", " ", CB_UF.Text)
                .Fields("CEP") = IIf(MEB_Cep.Text = "", " ", MEB_Cep.Text)
                Meb_Nasc.PromptInclude = False
                If Trim(Meb_Nasc.Text) = "" Then
                    Else
                        Meb_Nasc.PromptInclude = True
                        .Fields("Nasc") = Meb_Nasc.Text
                End If
                Meb_Nasc.PromptInclude = True
                .Update
        
        End Select
    End With
    Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina WHERE ProfID = " & ProfessorID & " ORDER BY ProfID")
    If RsProfessorDisciplina.BOF And RsProfessorDisciplina.EOF Then
        Else
            RsProfessorDisciplina.MoveFirst
            Do Until RsProfessorDisciplina.EOF
                RsProfessorDisciplina.Delete
                RsProfessorDisciplina.MoveNext
            Loop
    End If
    cont = 0
    Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina")
    Set RsDisciplinas = BD.OpenRecordset("SELECT * FROM Disciplina")
    Do Until cont = Lst_Disciplinas.ListCount
        If Lst_Disciplinas.Selected(cont) = True Then
            RsDisciplinas.FindFirst "Descr = '" & Lst_Disciplinas.List(cont) & "'"
            If RsDisciplinas.NoMatch Then
                    MsgBox "Erro ao consultar o banco de dados das Disciplinas. Por favor, verifique!", vbInformation, "CESNet - Aviso!"
                    Exit Sub
                Else
                    RsProfessorDisciplina.AddNew
                    RsProfessorDisciplina.Fields("ProfID") = ProfessorID
                    RsProfessorDisciplina.Fields("DisciplinaID") = RsDisciplinas.Fields("ID")
                    RsProfessorDisciplina.Update
            End If
        End If
        cont = cont + 1
    Loop
    

End Sub


'Private Sub Cb_Chv_KeyPress(KeyAscii As Integer)
'    KeyAscii = Asc(UCase(Chr(KeyAscii)))
' End Sub




Private Sub Cb_Nome_Click()
    If Cb_Nome.Text = "" Then Exit Sub
    With RsProfessor
        .FindFirst "Nome ='" & Cb_Nome.Text & "'"
        If .NoMatch Then
                MsgBox "Erro no acesso ao Banco de Dados." & Chr(13) & "Por favor, reinicie o formulário!", vbExclamation, "aviso!"
                Unload Me
            Else
                MstDadosProfessor
                LstDisciplinas
        End If
    End With
End Sub

Private Sub LstNomeProfessor()
    With RsProfessor
        If .BOF And .EOF Then
                Exit Sub
            Else
                .MoveFirst
                tmp = Cb_Nome.Text
                Cb_Nome.Clear
                Cb_Nome.Text = tmp
                Do While .EOF = False
                    Cb_Nome.AddItem (.Fields("Nome"))
                    .MoveNext
                Loop
        End If
    End With
End Sub
Private Sub MstDadosProfessor()
 With RsProfessor
        ProfessorID = .Fields("ProfID")
        UsuID = .Fields("UsuarioID")
        If UsuID = 0 Then
                MsgBox "Professor sem acesso ao sistema.", vbInformation, "CESNet - Atenção"
                UsuID = 0
            Else
                Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuID)
                If RsUsuario.BOF And RsUsuario.EOF Then
                        MsgBox "Erro ao localizar o Usuario.", vbInformation, "CESNet - Atenção"
                    Else
                        UsuID = RsUsuario.Fields("UsuarioID")
                End If
        End If
        If PgUsuNome(UsuID) = 0 Then
                Cb_Chv.Clear
            Else
                Cb_Chv.AddItem (PgUsuNome(UsuID))
                Cb_Chv.Text = PgUsuNome(UsuID)
        End If
        Cb_Nome.Text = .Fields("Nome")
        Txt_End.Text = IIf(IsNull(.Fields("End")), " ", .Fields("End"))
        Txt_Bai.Text = IIf(IsNull(.Fields("Bai")), " ", .Fields("Bai"))
        Txt_Mun.Text = IIf(IsNull(.Fields("Mun")), " ", .Fields("Mun"))
        CB_UF.Text = IIf(IsNull(.Fields("UF")), " ", .Fields("UF"))
        MEB_Cep.Text = IIf(IsNull(.Fields("CEP")), " ", .Fields("CEP"))
        Meb_Nasc.PromptInclude = False
        Meb_Nasc.Text = IIf(IsNull(.Fields("Nasc")), " ", .Fields("Nasc"))
        Meb_Nasc.PromptInclude = True
    End With
End Sub

Private Sub Cb_Nome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Acao = 0
    hdMenu (True)
    HDForm (False)

End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub hdMenu(op As Boolean)

    Tb_Menu.Buttons(1).Enabled = op
    Tb_Menu.Buttons(2).Enabled = op
    Tb_Menu.Buttons(3).Enabled = op
    Tb_Menu.Buttons(4).Enabled = op
    
    Tb_Menu.Buttons(6).Enabled = IIf(op = False, True, False)
    Tb_Menu.Buttons(7).Enabled = IIf(op = False, True, False)
    
End Sub


Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            Acao = 1
            HDForm (True)
            hdMenu (False)
            ProfessorID = ""
            LimpForm
            LstDisciplinas
            
        Case 2 'Alterar
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            Acao = 2
            
            If Cb_Nome.Text = "" Then Exit Sub
            
            If Cb_Chv.Text = Usuario Then
                MsgBox "O Usuario não pode alterar suas próprias informações.", vbInformation, "CESNet - Aviso!"
                Exit Sub
            End If
            HDForm (True)
            hdMenu (False)
            
        Case 3 'Excluir
            Acao = 3
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            If Cb_Nome.Text = "" Then
                Exit Sub
            End If
            tmp = MsgBox("Deseja excluir este cadastro?", vbYesNo, "CESNet - Aviso")
            If tmp = 6 Then
                RsProfessor.FindFirst "ProfID = " & ProfessorID
                RsProfessor.Delete
                BD.Execute "DELETE * FROM ProfessorDisciplina WHERE ProfID = " & ProfessorID
                LimpForm
            End If

            
        Case 4 'Imprimir
            Acao = 4
            'If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            'ImprimirListagem
        Case 6 'Gravar
            If ValidarSoftware("Professores") = False Then Exit Sub
            GrvProf
             Acao = 0
            HDForm (False)
            hdMenu (True)
            'LimpForm

        Case 7 'Cancelar
            Acao = 0
            HDForm (False)
            hdMenu (True)
            LimpForm
    
    End Select
End Sub

Private Sub LstDisciplinas()
    Set RsDisciplinas = BD.OpenRecordset("SELECT * FROM Disciplina")
    If RsDisciplinas.BOF And RsDisciplinas.EOF Then
        MsgBox "Nenhuma Disciplina cadastrada. Por Favor Verifique", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    With RsDisciplinas
        Lst_Disciplinas.Clear
        .MoveFirst
        Do Until .EOF
            Lst_Disciplinas.AddItem (.Fields("Descr"))
            .MoveNext
        Loop
        cont = 0
        If Acao = 1 Then Exit Sub
        Do Until cont = Lst_Disciplinas.ListCount
            Set RsDisciplinas = BD.OpenRecordset("SELECT * FROM Disciplina WHERE Descr = '" & Lst_Disciplinas.List(cont) & "'")
            Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina WHERE ProfID = " & ProfessorID & " AND DisciplinaID = " & RsDisciplinas.Fields("ID"))
            If RsProfessorDisciplina.BOF And RsProfessorDisciplina.EOF Then
                Else
                    Lst_Disciplinas.Selected(cont) = True
            End If
            cont = cont + 1
        Loop
    End With
End Sub
Private Sub Txt_Bai_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_End_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub Txt_Mun_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub
Private Function PgUsuIDID(ProfID As String)
    If Trim(ProfID) = "" Then
        PgUsuIDID = ""
        Exit Function
    End If
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE Chv = '" & ProfID & "'")
    If RsUsuario.BOF And RsUsuario.EOF Then
            'MsgBox "Usuario não cadastrado.", vbInformation ,"CESNet - Aviso!"
            PgUsuIDID = 0
            Cb_Chv.SetFocus
            Exit Function
        Else
            PgUsuIDID = RsUsuario.Fields("UsuarioID")
    End If
End Function
Private Function PgUsuNome(ID As Integer)
    If Trim(ID) = "" Then
        PgUsuNome = ""
        Exit Function
    End If
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & ID)
    If RsUsuario.BOF And RsUsuario.EOF Then
            'MsgBox "Usuario não cadastrado.", vbInformation ,"CESNet - Aviso!"
            PgUsuNome = 0
            'Cb_Chv.SetFocus
            Exit Function
        Else
            PgUsuNome = RsUsuario.Fields("Chv")
    End If
End Function

