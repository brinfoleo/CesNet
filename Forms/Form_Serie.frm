VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Serie 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Série"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5520
   Icon            =   "Form_Serie.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   5520
   Begin VB.TextBox txtSigla 
      Height          =   285
      Left            =   4500
      MaxLength       =   4
      TabIndex        =   4
      Top             =   1200
      Width           =   615
   End
   Begin VB.TextBox txtDescr 
      Enabled         =   0   'False
      Height          =   285
      Left            =   180
      MaxLength       =   20
      TabIndex        =   0
      Top             =   1200
      Width           =   3780
   End
   Begin MSFlexGridLib.MSFlexGrid MSFGDescr 
      Height          =   3645
      Left            =   60
      TabIndex        =   1
      Top             =   1560
      Width           =   5385
      _ExtentX        =   9499
      _ExtentY        =   6429
      _Version        =   393216
      Cols            =   3
      FormatString    =   "^ID   |^SIGLA |^DESCRIÇÃO                                                               "
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   2
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
               Picture         =   "Form_Serie.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":0624
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":093E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":0C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":0F72
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":128C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":15A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":18C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":1BDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":1EF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":220E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":2528
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":2842
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":2B5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":2E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":3190
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":34AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":37C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":3ADE
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":3DF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":4112
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":442C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":4746
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":4A60
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":4D7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":5094
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":53AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":56C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":59E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":5CFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":6016
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Serie.frx":6330
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CADASTRO DE SÉRIE"
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
      TabIndex        =   6
      Top             =   600
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "Sigla:"
      Height          =   195
      Left            =   4500
      TabIndex        =   5
      Top             =   960
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Série:"
      Height          =   195
      Left            =   180
      TabIndex        =   3
      Top             =   960
      Width           =   855
   End
End
Attribute VB_Name = "Form_Serie"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim ID      As Integer
Dim Tabela  As String
Dim Acao    As Integer

Private Sub Form_Load()
    Tabela = "Serie"
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
            ID = MSFGDescr.TextMatrix(MSFGDescr.Row, 0)
            txtSigla.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 1)
            txtDescr.Text = MSFGDescr.TextMatrix(MSFGDescr.Row, 2)
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
            CabImp7 ("MÓDULO")
            Printer.FontSize = 10
            Printer.FontBold = False
            Printer.FontItalic = False
            Printer.FontUnderline = False
            Printer.Print Tab(20); String(100, "-")
            Printer.Print Tab(25); "ID", Tab(35); "SIGLA"; Tab(45); "DESCRIÇÃO"
            Printer.Print Tab(20); String(100, "-")
            Do Until RsImpr.EOF
                Printer.Print Tab(25); left("00", 2 - Len(Trim(RsImpr.Fields("ID")))) & RsImpr.Fields("ID"); Tab(35); IIf(IsNull(RsImpr.Fields("Sigla")), " ", Trim(RsImpr.Fields("Sigla"))); Tab(45); RsImpr.Fields("Descr")
                RsImpr.MoveNext
            Loop
            Printer.EndDoc
            RsImpr.Close
    End If
End Sub
Private Sub ExcluirCampo()
    If ID = 0 Then Exit Sub
    If MsgBox("Deseja realmente EXCLUIR o item " & ID & "?", vbInformation + vbYesNo, "Exclusão") = vbYes Then
        BD.Execute "DELETE * FROM " & Tabela & " WHERE ID = " & ID
        RegLog "", "Serie (" & ID & ")" & PgNomeSerie(ID) & " excluida"
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
    
    If ID = 0 Then
            Set RsTMP = BD.OpenRecordset("SELECT * FROM " & Tabela)
            RsTMP.AddNew
        Else
            Set RsTMP = BD.OpenRecordset("SELECT * FROM " & Tabela & " WHERE ID = " & ID)
            If RsTMP.BOF And RsTMP.EOF Then
                    MsgBox "Erro ao gravar dados", vbCritical, "Aviso"
                    RsTMP.Close
                    Exit Sub
                Else
                    RsTMP.Edit
            End If
    End If
    
    RsTMP.Fields("Descr") = Trim(txtDescr.Text)
    RsTMP.Fields("Sigla") = Trim(txtSigla.Text)
    RsTMP.Update
    RsTMP.Close
End Sub
Private Sub HDForm(op As Boolean)
    txtDescr.Enabled = op
    txtSigla.Enabled = op
    MSFGDescr.Enabled = IIf(op = True, False, True)
End Sub
Private Sub LimpForm()
    ID = 0
    txtDescr.Text = ""
    txtSigla.Text = ""
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
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 1) = IIf(IsNull(RsTMP.Fields("Sigla")), "", RsTMP.Fields("Sigla"))
                MSFGDescr.TextMatrix(MSFGDescr.Rows - 1, 2) = RsTMP.Fields("Descr")
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


Private Sub txtDescr_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

Private Sub txtSigla_KeyPress(KeyAscii As Integer)
        KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub
