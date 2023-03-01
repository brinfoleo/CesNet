VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_InstEnsino 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Cadastro de Inst. Ensino"
   ClientHeight    =   6795
   ClientLeft      =   90
   ClientTop       =   345
   ClientWidth     =   11850
   Icon            =   "Form_InstEnsino.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6795
   ScaleWidth      =   11850
   Begin VB.TextBox txtCidadeRed 
      Height          =   285
      Left            =   7800
      MaxLength       =   30
      TabIndex        =   11
      Top             =   1380
      Width           =   3495
   End
   Begin VB.ComboBox cbUF 
      Enabled         =   0   'False
      Height          =   315
      ItemData        =   "Form_InstEnsino.frx":030A
      Left            =   1500
      List            =   "Form_InstEnsino.frx":035F
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   1740
      Width           =   750
   End
   Begin VB.TextBox txtCidade 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1500
      MaxLength       =   30
      TabIndex        =   6
      Top             =   1380
      Width           =   4140
   End
   Begin VB.TextBox txtSigla 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7800
      MaxLength       =   20
      TabIndex        =   3
      Top             =   1020
      Width           =   1260
   End
   Begin VB.TextBox txtDescr 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1500
      MaxLength       =   100
      TabIndex        =   0
      Top             =   1020
      Width           =   5265
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   11835
      _ExtentX        =   20876
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
               Picture         =   "Form_InstEnsino.frx":03CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":06E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":0A02
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":0D1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":1036
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":1350
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":166A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":1984
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":1C9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":1FB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":22D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":25EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":2906
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":2C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":2F3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":3254
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":356E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":3888
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":3BA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":3EBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":41D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":44F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":480A
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":4B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":4E3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":5158
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":5472
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":578C
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":5AA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":5DC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":60DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_InstEnsino.frx":63F4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFGDescr 
      Height          =   4590
      Left            =   60
      TabIndex        =   10
      Top             =   2160
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   8096
      _Version        =   393216
      Cols            =   6
      FixedCols       =   0
      ScrollBars      =   2
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"Form_InstEnsino.frx":670E
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Cidade(Reduzida):"
      Height          =   195
      Left            =   6180
      TabIndex        =   12
      Top             =   1440
      Width           =   1575
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CADASTRO DE INSTITUIÇÃO DE ENSINO"
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
      TabIndex        =   8
      Top             =   600
      Width           =   11835
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "UF:"
      Height          =   240
      Left            =   780
      TabIndex        =   5
      Top             =   1800
      Width           =   660
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Cidade:"
      Height          =   195
      Left            =   165
      TabIndex        =   4
      Top             =   1440
      Width           =   1290
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Sigla:"
      Height          =   195
      Left            =   7170
      TabIndex        =   2
      Top             =   1065
      Width           =   585
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Inst. de Ensino:"
      Height          =   210
      Left            =   30
      TabIndex        =   1
      Top             =   1065
      Width           =   1425
   End
End
Attribute VB_Name = "Form_InstEnsino"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim id      As Integer
Dim Tabela  As String
Dim Acao    As Integer


Private Sub Form_Load()
    Tabela = "InstEnsino"
    HDForm (False)
    hdMenu (True)
    MstDados
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



Private Sub MSFGDescr_Click()
    If MSFGDescr.MouseRow = 0 Then
            LimpForm
            Exit Sub
        Else
            id = MSFGDescr.TextMatrix(MSFGDescr.Row, 0)
            txtDescr.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 1)
            txtSigla.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 2)
            txtCidade.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 3)
            txtCidadeRed.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 4)
            cbUF.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 5)
    End If
    
    
End Sub

Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            Acao = 1
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            HDForm (True)
            hdMenu (False)
            LimpForm
        Case 2 'Alterar
            Acao = 2
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            If Trim(txtDescr.Text) = "" Then Exit Sub
            HDForm (True)
            hdMenu (False)
        Case 3 'Excluir
            Acao = 3
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            'LimpForm
            ExcluirCampo
            
        Case 4 'Imprimir
            Acao = 4
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            ImprimirListagem
        Case 6 'Gravar
            If ValidarSoftware(Tabela) = False Then Exit Sub
            GravarDados
            LimpForm
            HDForm (False)
            hdMenu (True)
            MstDados
        Case 7 'Cancelar
            Acao = 7
            HDForm (False)
            hdMenu (True)
            LimpForm
    End Select
End Sub
Private Sub ImprimirListagem()
    Dim RsImpr As Recordset
    
    Set RsImpr = BD.OpenRecordset("SELECT * FROM " & Tabela & " ORDER BY ID")
    If RsImpr.BOF And RsImpr.EOF Then
            MsgBox "Operação cancelada! Não existe registros cadastrados.", vbInformation, "Aviso"
            RsImpr.Close
        Else
            RsImpr.MoveFirst
            CabImp7 ("INSTITUIÇÃO DE ENSINO")
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            
            Printer.Print Tab(5); String(170, "-")
            Printer.Print Tab(5); "ID"; Tab(11); "SIGLA";
            Printer.Print Tab(11); "DESCRIÇÃO"; Tab(85); "CIDADE / UF"
            Printer.Print Tab(5); String(170, "-")
            
            Do Until RsImpr.EOF
                
                Printer.Print Tab(5); left("00", 2 - Len(Trim(RsImpr.Fields("ID")))) & RsImpr.Fields("ID"); Tab(11); RsImpr.Fields("Sigla")
                Printer.Print Tab(11); RsImpr.Fields("Descr"); Tab(85); RsImpr.Fields("Cidade") & "(" & RsImpr.Fields("CidadeRed") & ") - " & RsImpr.Fields("UF")
                Printer.Print
                RsImpr.MoveNext
            Loop
            Printer.EndDoc
            RsImpr.Close
    End If
End Sub
Private Sub ExcluirCampo()
    If id = 0 Then Exit Sub
    If MsgBox("Deseja realmente EXCLUIR o item " & id & "?", vbInformation + vbYesNo, "Exclusão") = vbYes Then
        BD.Execute "DELETE * FROM " & Tabela & " WHERE ID = " & id
        MsgBox "Item EXCLUIDO com sucesso!", vbInformation, "Aviso"
        LimpForm
        MstDados
    End If
End Sub
Private Sub GravarDados()
    Dim RsTMP As Recordset
    
    If txtDescr.Text = "" Then
        MsgBox "Operação cancelada devido não pode gravada uma informação nula.", vbInformation, "Aviso"
        Exit Sub
    End If
    If ExisteReg = True Then
        MsgBox "Operação cancelada devido haver ambiguidade de nomes.", vbInformation, "Aviso"
        Exit Sub
    End If
    
    If id = 0 Then
            Set RsTMP = BD.OpenRecordset("SELECT * FROM " & Tabela)
            RsTMP.AddNew
        Else
            Set RsTMP = BD.OpenRecordset("SELECT * FROM " & Tabela & " WHERE ID = " & id)
            If RsTMP.BOF And RsTMP.EOF Then
                    MsgBox "Erro ao gravar dados", vbCritical, "Aviso"
                    RsTMP.Close
                    Exit Sub
                Else
                    RsTMP.Edit
            End If
    End If
    
    RsTMP.Fields("Descr") = Trim(txtDescr.Text)
    RsTMP.Fields("Sigla") = IIf(Trim(txtSigla.Text) = "", Null, Trim(txtSigla.Text))
    RsTMP.Fields("Cidade") = Trim(txtCidade.Text)
    RsTMP.Fields("CidadeRed") = Trim(txtCidadeRed.Text)
    RsTMP.Fields("UF") = Trim(cbUF.Text)
    RsTMP.Update
    RsTMP.Close
End Sub
Private Sub HDForm(op As Boolean)
    txtDescr.Enabled = op
    txtSigla.Enabled = op
    txtCidade.Enabled = op
    txtCidadeRed.Enabled = op
    cbUF.Enabled = op
    MSFGDescr.Enabled = IIf(op = True, False, True)
End Sub
Private Sub LimpForm()
    id = 0
    txtDescr.Text = ""
    txtSigla.Text = ""
    txtCidade.Text = ""
    txtCidadeRed.Text = ""
    cbUF.Text = " "
End Sub
Private Sub MstDados()
    Dim RsTMP As Recordset
    MSFGDescr.Rows = 1
    Set RsTMP = BD.OpenRecordset("SELECT * FROM " & Tabela)
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                MSFGDescr.Rows = MSFGDescr.Rows + 1
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 0) = RsTMP.Fields("ID")
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 1) = RsTMP.Fields("Descr")
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 2) = IIf(IsNull(RsTMP.Fields("Sigla")), "", RsTMP.Fields("Sigla"))
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 3) = IIf(IsNull(RsTMP.Fields("Cidade")), "", RsTMP.Fields("Cidade"))
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 4) = IIf(IsNull(RsTMP.Fields("CidadeRed")), "", RsTMP.Fields("CidadeRed"))
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 5) = IIf(IsNull(RsTMP.Fields("UF")), "", RsTMP.Fields("UF"))
                RsTMP.MoveNext
            Loop
            RsTMP.Close
    End If
End Sub
Private Function ExisteReg() As Boolean
'True- Existe // False - Nao Exite
    If Acao = 2 Then
        ExisteReg = False
        Exit Function
    End If
    Dim RsTMP As Recordset
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM " & Tabela & " WHERE Descr = '" & Trim(txtDescr.Text) & "'")
    If RsTMP.BOF And RsTMP.EOF Then
            ExisteReg = False
        Else
            ExisteReg = True
    End If
    RsTMP.Close
End Function


Private Sub txtCidade_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub txtCidadeRed_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub


Private Sub txtDescr_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtSigla_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

