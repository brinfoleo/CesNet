VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_Usuario 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Usuário"
   ClientHeight    =   6765
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8640
   Icon            =   "Form_Usuario.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6765
   ScaleWidth      =   8640
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   5715
      TabIndex        =   19
      Top             =   6075
      Width           =   2670
   End
   Begin VB.CommandButton Bt_AltUsu 
      Caption         =   "Alterar Usuário"
      Height          =   555
      Left            =   5715
      TabIndex        =   18
      Top             =   3735
      Width           =   2670
   End
   Begin VB.CommandButton Bt_NovoUsu 
      Caption         =   "Novo Usuário"
      Height          =   555
      Left            =   5715
      TabIndex        =   17
      Top             =   3150
      Width           =   2670
   End
   Begin VB.Frame Frame2 
      Caption         =   "Usuário:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2085
      Left            =   90
      TabIndex        =   8
      Top             =   3060
      Width           =   5325
      Begin VB.ComboBox Cb_Grupo 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   1350
         Width           =   3660
      End
      Begin VB.TextBox Txt_Resp 
         Enabled         =   0   'False
         Height          =   330
         Left            =   1260
         MaxLength       =   50
         TabIndex        =   12
         Top             =   810
         Width           =   3615
      End
      Begin VB.TextBox Txt_Chv 
         Enabled         =   0   'False
         Height          =   315
         Left            =   1260
         MaxLength       =   10
         TabIndex        =   9
         Top             =   315
         Width           =   1875
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Grupo:"
         Height          =   195
         Left            =   630
         TabIndex        =   13
         Top             =   1440
         Width           =   465
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Responsavel:"
         Height          =   195
         Left            =   135
         TabIndex        =   11
         Top             =   900
         Width           =   960
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Login:"
         Height          =   195
         Left            =   675
         TabIndex        =   10
         Top             =   405
         Width           =   420
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Usuários:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2535
      Left            =   90
      TabIndex        =   6
      Top             =   405
      Width           =   8475
      Begin MSFlexGridLib.MSFlexGrid MSFG_Usuarios 
         Height          =   2175
         Left            =   180
         TabIndex        =   7
         Top             =   225
         Width           =   8205
         _ExtentX        =   14473
         _ExtentY        =   3836
         _Version        =   393216
         Cols            =   7
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_Usuario.frx":030A
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Opções:"
      Height          =   1455
      Left            =   90
      TabIndex        =   3
      Top             =   5175
      Width           =   5325
      Begin VB.CheckBox Chk_PwsPadrao 
         Caption         =   "Senha padrão: 123"
         Enabled         =   0   'False
         Height          =   240
         Left            =   2655
         TabIndex        =   20
         Top             =   1125
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.TextBox Txt_SenhaIni 
         Enabled         =   0   'False
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1170
         MaxLength       =   6
         PasswordChar    =   "#"
         TabIndex        =   16
         Text            =   "123"
         Top             =   1080
         Width           =   1320
      End
      Begin VB.CheckBox Chk_Opcoes 
         Caption         =   "Senha nunca expira."
         Enabled         =   0   'False
         Height          =   195
         Index           =   1
         Left            =   180
         TabIndex        =   5
         Top             =   720
         Width           =   3030
      End
      Begin VB.CheckBox Chk_Opcoes 
         Caption         =   "Trocar senha no proximo acesso."
         Enabled         =   0   'False
         Height          =   195
         Index           =   0
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   3030
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Senha inicial:"
         Height          =   195
         Left            =   135
         TabIndex        =   15
         Top             =   1125
         Width           =   960
      End
   End
   Begin VB.CommandButton Bt_ExcUsu 
      Caption         =   "Excluir Usuário"
      Height          =   555
      Left            =   5715
      TabIndex        =   1
      Top             =   4320
      Width           =   2685
   End
   Begin VB.CommandButton Bt_GravUsuario 
      Caption         =   "Gravar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   5715
      TabIndex        =   0
      Top             =   5490
      Width           =   2685
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "USUÁRIO"
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
      TabIndex        =   2
      Top             =   0
      Width           =   8595
   End
End
Attribute VB_Name = "Form_Usuario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUsuario As Recordset
Dim tmp As Integer
Dim Senha As String
Dim UsuID As Integer
Dim opcao As Integer  '1 - Novo   // 2 = Alterar




Private Sub Bt_AltUsu_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    HDForm (True)
    
    opcao = 2
End Sub

Private Sub Bt_Cancelar_Click()
    LimpForm
    HDForm (False)
End Sub

Private Sub Bt_ExcUsu_Click()
    If ChkAcesso(Me.Name, "E") = False Then Exit Sub
    If MsgBox("Deseja excluir o usuário: " & Txt_Chv.Text, vbInformation + vbYesNo, "CESNet - Aviso!") = vbYes Then
        'MsgBox "Os campos RESPONSAVEL e CHAVE nao podem ser em branco!", vbInformation, "CESNet - Aviso!"
        'Exit Sub
        BD.Execute "DELETE * FROM Usuario WHERE UsuarioID = " & UsuID
    End If
    'Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario")
    'RsUsuario.FindFirst "Responsavel = '" & Txt_Resp.Text & "' OR Chv = '" & Txt_Chv.Text & "'"
    'With RsUsuario
    '    If .NoMatch Then
    '            MsgBox "Nenhum Usuario encontrado.", vbInformation, "CESNet - Aviso!"
    '            Exit Sub
    '        Else
    '            .Delete
    '    End If
    'End With
    LimpForm
    LstUsuarios
    
End Sub

Private Sub Bt_GravUsuario_Click()
    If Trim(Cb_Grupo.Text) = "" Then
        MsgBox "O campo GRUPO não pode ser em branco!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    If Trim(Txt_Resp.Text) = "" Then
        MsgBox "O campo RESPONSAVEL não pode ser em branco!", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    If Chk_Opcoes(0).Value = 1 Then
        If Trim(Txt_SenhaIni.Text) = "" Then
            MsgBox "O campo SENHA INICIAL não pode ser em branco!", vbInformation, "CESNet - Aviso!"
            Exit Sub
        End If
    End If
    
    If Len(Trim(Txt_Chv.Text)) < 3 Then
            MsgBox "A CHAVE não pode ter menos de 3 caracteres. Por favor Verifique!", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            For tmp = 1 To Len(Txt_Chv.Text)
                If Mid(Txt_Chv.Text, tmp, 1) = " " Then
                    MsgBox "O campo CHAVE não pode conter espaços em branco. Por favor verifique!", vbInformation, "CESNet - Aviso!"
                    Txt_Chv.SetFocus
                    Exit Sub
                End If
            Next
    End If
    
    If ValidarSoftware("Usuario") = False Then Exit Sub
    
    Select Case opcao
        Case 1
            Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE Chv = '" & Trim(Txt_Chv.Text) & "'")
            With RsUsuario
            
                If .BOF And .EOF Then
                        .AddNew
                        .Fields("Responsavel") = Trim(Txt_Resp.Text)
                        .Fields("Chv") = Trim(Txt_Chv.Text)
                        .Fields("Pwd") = Crypto(Txt_SenhaIni.Text)
                        .Fields("GrupoID") = PgIDGrupo(Cb_Grupo.Text)
                        
                        .Fields("Acesso") = 0
                        .Fields("DtInclusao") = Date
                        .Fields("TrocarPWD") = IIf(Chk_Opcoes(0).Value = 1, True, False)
                        .Fields("PWDNuncaExp") = IIf(Chk_Opcoes(1).Value = 1, True, False)
                        .Update
                    Else
                        MsgBox "Chave Invalida. Por favor escolha outra chave!", vbInformation, "CESNet - Aviso!"
                        Txt_Chv.SetFocus
                        Exit Sub
                End If
            End With
        Case 2
            Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuID)
            With RsUsuario
                .MoveFirst
                .Edit
                .Fields("Responsavel") = Trim(Txt_Resp.Text)
                .Fields("Chv") = Trim(Txt_Chv.Text)
                '.Fields("Pwd") = Crypto(Txt_SenhaIni.Text)
                .Fields("GrupoID") = PgIDGrupo(Cb_Grupo.Text)
                        
                '.Fields("Acesso") = 0
                '.Fields("DtInclusao") = Date
                
                .Fields("TrocarPWD") = IIf(Chk_Opcoes(0).Value = 1, True, False)
                .Fields("PWDNuncaExp") = IIf(Chk_Opcoes(1).Value = 1, True, False)
                If Chk_Opcoes(0).Value = 1 Then
                    .Fields("Pwd") = Crypto(Txt_SenhaIni.Text)
                End If
                .Update
            End With
                '.Fields("TrocarPWD") = True
                '.Update
        End Select
        
   
    
'    Dim strAtivacao As String  'Variavel que contem 0 e 1
'    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario WHERE Responsavel = '" & txt_resp.Text & "' AND Chv = '" & Txt_Chv.Text & "'")
'        If RsUsuario.BOF And RsUsuario.EOF Then
'            MsgBox "Erro no acesso ao tente novamente", vbInformation, "CESNet - Aviso!"
'            Exit Sub
'        End If
 '       RsUsuario.MoveFirst
 '
'        For Tmp = 1 To TrVw_Sistema.Nodes.Count
'
'            If TrVw_Sistema.Nodes.Item(Tmp).Checked = True Then
'                    strAtivacao = strAtivacao & "1"
'                Else
'                    strAtivacao = strAtivacao & "0"
'            End If
'            ''Debug.Print Tmp & " - " & TrVw_Sistema.Nodes.Item(Tmp)
'        Next
'        RsUsuario.Edit
'        RsUsuario.Fields("Sessao") = strAtivacao
'        RsUsuario.Update
'    LimpForm
    
    'Usuario = Trim(Txt_Chv.Text)
    'UsuarioID = RsUsuario.Fields("UsuarioID")
    
    'Permicoes (UsuarioID)
    'MDIForm_Main.LoadStatusBarr
    HDForm (False)
    LimpForm
    LstUsuarios
End Sub








Private Sub Bt_NovoUsus_Click()

    
End Sub

Private Sub Bt_NovoUsu_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    LimpForm
    HDForm (True)
    opcao = 1
End Sub


Private Sub Cb_Grupo_DropDown()
    Cb_Grupo.Clear
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM UsuGrupo ORDER BY Nome ASC")
    If RsUsuario.BOF And RsUsuario.EOF Then
            
        Else
            RsUsuario.MoveFirst
            Do Until RsUsuario.EOF
                Cb_Grupo.AddItem RsUsuario.Fields("Nome")
                RsUsuario.MoveNext
            Loop
    End If
    RsUsuario.Close
End Sub


Private Sub Chk_Opcoes_Click(Index As Integer)
    If Chk_Opcoes(0).Value = 1 Then
        Txt_SenhaIni.Text = ""
        Txt_SenhaIni.Enabled = True
        Chk_PwsPadrao.Value = 0
        Chk_PwsPadrao.Enabled = True
        
        
    End If
End Sub

Private Sub Chk_PwsPadrao_Click()
    If Chk_PwsPadrao.Value = 0 Then
            Txt_SenhaIni.Text = ""
            Txt_SenhaIni.Enabled = True
        Else
            Txt_SenhaIni.Text = "123"
            Txt_SenhaIni.Enabled = False
    End If
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    LstUsuarios
    opcao = 0
End Sub


Private Sub MSFG_Usuarios_Click()
    opcao = 2
    UsuID = MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 0)
    Txt_Chv.Text = MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 2)
    Txt_Resp.Text = MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 1)
    Cb_Grupo.AddItem MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 3)
    Cb_Grupo.Text = MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 3)
    Chk_Opcoes(0).Value = IIf(MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 5) = "SIM", 1, 0)
    Chk_Opcoes(1).Value = IIf(MSFG_Usuarios.TextMatrix(MSFG_Usuarios.Row, 6) = "SIM", 1, 0)
End Sub

Private Sub Txt_Chv_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Or KeyAscii = 0 Then Exit Sub
    If KeyAscii >= 65 And KeyAscii <= 90 Then 'Letras Maiusculas
            Else
                If KeyAscii >= 48 And KeyAscii <= 57 Then 'Numeros entre 0 a 9
                        Exit Sub
                    Else
                        KeyAscii = 0
                        Beep
                End If
    End If
End Sub
Private Sub LimpForm()
    Cb_Grupo.Clear
    Txt_Chv.Text = ""
    Txt_Resp.Text = ""
End Sub
Private Sub LstUsuarios()
    
    MSFG_Usuarios.Rows = 1
    Set RsUsuario = BD.OpenRecordset("SELECT * FROM Usuario ORDER BY Responsavel")
    If RsUsuario.BOF And RsUsuario.EOF Then
            RsUsuario.Close
        Else
            RsUsuario.MoveFirst
            Do Until RsUsuario.EOF
                With MSFG_Usuarios
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = RsUsuario.Fields("UsuarioID")
                    .TextMatrix(.Rows - 1, 1) = RsUsuario.Fields("Responsavel")
                    .TextMatrix(.Rows - 1, 2) = RsUsuario.Fields("Chv") '& Crypto(RsUsuario.Fields("PWD"))
                    .TextMatrix(.Rows - 1, 3) = PgNomeGrupo(IIf(IsNull(RsUsuario.Fields("GrupoID")), "0", RsUsuario.Fields("GrupoID")))
                    .TextMatrix(.Rows - 1, 4) = RsUsuario.Fields("Acesso")
                    .TextMatrix(.Rows - 1, 5) = IIf(RsUsuario.Fields("TrocarPWD") = True, "SIM", "NÃO")
                    .TextMatrix(.Rows - 1, 6) = IIf(RsUsuario.Fields("PWDNuncaExp") = True, "SIM", "NÃO")
                    RsUsuario.MoveNext
                End With
            Loop
    End If
    RsUsuario.Close
End Sub

Private Sub LstAcesso()
    Txt_Chv.Text = RsUsuario.Fields("Chv")
    Dim Grupo As String  'Variavel que contem 0 e 1
    Grupo = IIf(IsNull(RsUsuario.Fields("GrupoID")), "", RsUsuario.Fields("GrupoID"))
    
End Sub
Private Sub HDForm(op As Boolean)
    'Bt_GravUsuario.Enabled = op
    'Bt_ExcUsu.Enabled = op
    MSFG_Usuarios.Enabled = IIf(op = True, False, True)
    Bt_NovoUsu.Enabled = IIf(op = True, False, True)
    Bt_ExcUsu.Enabled = IIf(op = True, False, True)
    Bt_AltUsu.Enabled = IIf(op = True, False, True)
    
    Bt_GravUsuario.Enabled = op
    Bt_Cancelar.Enabled = op
    
    Txt_Chv.Enabled = op
    Txt_Resp.Enabled = op
    Cb_Grupo.Enabled = op
    
    Chk_Opcoes(0).Enabled = op
    Chk_Opcoes(1).Enabled = op
    
    Chk_PwsPadrao.Enabled = op
    'Txt_SenhaIni.Enabled = op
    'Chk_PwsPadrao.Enabled = op
    
End Sub

Private Sub Txt_Resp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


