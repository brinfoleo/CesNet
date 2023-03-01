VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form_Mail 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Correio Interno"
   ClientHeight    =   7305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9870
   Icon            =   "Form_Mail.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7305
   ScaleWidth      =   9870
   Begin VB.Frame Frame2 
      Height          =   2535
      Left            =   60
      TabIndex        =   9
      Top             =   960
      Width           =   9735
      Begin MSComctlLib.ListView ListView1 
         Height          =   2055
         Left            =   60
         TabIndex        =   10
         Top             =   180
         Width           =   9615
         _ExtentX        =   16960
         _ExtentY        =   3625
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         Icons           =   "ilMail"
         SmallIcons      =   "ilMail"
         ColHdrIcons     =   "ilMail"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   4
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "ID"
            Object.Width           =   1058
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Data"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "De"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Assunto"
            Object.Width           =   10583
         EndProperty
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Duplo click para marcar como NOVO ou LIDO."
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   2280
         Width           =   6915
      End
   End
   Begin VB.Frame Frame1 
      Height          =   3675
      Left            =   60
      TabIndex        =   1
      Top             =   3540
      Width           =   9735
      Begin VB.TextBox txtDescr 
         Height          =   2355
         Left            =   60
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   1200
         Width           =   9495
      End
      Begin VB.Label lbAssunto 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   7
         Top             =   840
         Width           =   8475
      End
      Begin VB.Label lbDe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   540
         Width           =   3675
      End
      Begin VB.Label lbData 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   960
         TabIndex        =   5
         Top             =   240
         Width           =   3675
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Assunto:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   840
         Width           =   795
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "De:"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   540
         Width           =   675
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   255
         Left            =   420
         TabIndex        =   2
         Top             =   240
         Width           =   495
      End
   End
   Begin MSComctlLib.ImageList ilMail 
      Left            =   7440
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Mail.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form_Mail.frx":0BE4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1058
      ButtonWidth     =   1191
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "IL_Menu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   8
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Alterar"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Gravar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Object.Visible         =   0   'False
            Caption         =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Atualiza"
            ImageIndex      =   33
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
            NumListImages   =   33
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":14BE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":17D8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":1AF2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":1E0C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":2126
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":2440
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":275A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":2A74
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":2D8E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":30A8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":33C2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":36DC
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":39F6
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":3D10
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":402A
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":4344
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":465E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":4978
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":4C92
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":4FAC
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":52C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":55E0
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":58FA
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":5C14
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":5F2E
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":6248
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":6562
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":687C
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":6B96
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":6EB0
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":71CA
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":74E4
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Mail.frx":77FE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CORREIO INTERNO"
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
      TabIndex        =   12
      Top             =   600
      Width           =   9885
   End
End
Attribute VB_Name = "Form_Mail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MailID As Integer

Private Sub HDForm(op As Boolean)

    lbData.Enabled = op
    lbDe.Enabled = op
    lbAssunto.Enabled = op
    txtDescr.Enabled = op

End Sub

Private Sub LimpForm()
    lbData.Caption = ""
    lbDe.Caption = ""
    lbAssunto.Caption = ""
    txtDescr.Text = ""
    MailID = 0
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
    hdMenu (True)
    ExibirGrid
End Sub

Private Sub ExibirGrid()
    Dim RsMail      As Recordset
    Dim itmx        As ListItem
    Dim colx        As ColumnHeader
    Dim codImg      As Integer
    
    ListView1.ListItems.Clear
    
    
    'Dim i           As Integer
    'Inclui algumas colunas
    'Set colx = ListView1.ColumnHeaders.Add(, , "Nome")
    'Set colx = ListView1.ColumnHeaders.Add(, , "Tipo")
    'Set colx = ListView1.ColumnHeaders.Add(, , "Tam.")
    'Set colx = ListView1.ColumnHeaders.Add(, , "Data")
    Set RsMail = BD.OpenRecordset("SELECT * FROM Mail WHERE Para = " & UsuarioID & " ORDER BY Data, NOVO")
    If RsMail.BOF And RsMail.EOF Then
        Else
            RsMail.MoveFirst
            Do Until RsMail.EOF
                codImg = IIf(RsMail.Fields("Novo") = True, 2, 1)
                Set itmx = ListView1.ListItems.Add(, , RsMail.Fields("ID"), codImg, codImg)
                'Aqui estamos acessando e definindo cada subitem
                itmx.SubItems(1) = IIf(IsNull(RsMail.Fields("Data")), " ", RsMail.Fields("Data"))
                itmx.SubItems(2) = PgRespUsu(RsMail.Fields("DE"))
                itmx.SubItems(3) = IIf(IsNull(RsMail.Fields("Assunto")), " ", RsMail.Fields("Assunto"))
                'itmx.SubItems(3) = "01/04/2001"
                'Define o formato de visao como Report
                        'ListView1.View = lvwReport
                RsMail.MoveNext
        Loop
    End If

End Sub





Private Sub ListView1_Click()
    On Error GoTo TtERRO
    
    
    
    MailID = ListView1.SelectedItem.Text
    
    
    
    
    ExibirMail (MailID)
    Exit Sub
TtERRO:
    Exit Sub
End Sub

Private Sub ListView1_DblClick()
    Dim imgMail     As Integer
    Dim RsMail      As Recordset
    'Dim MailID      As Integer
    
    'MailID = ListView1.SelectedItem.Text

    Set RsMail = BD.OpenRecordset("SELECT * FROM Mail WHERE ID = " & MailID)
    If RsMail.BOF And RsMail.EOF Then
            MsgBox "Erro ao localizar Mail", vbInformation, "CESNet - Aviso"
        Else
            RsMail.MoveFirst
            RsMail.Edit
            RsMail.Fields("Novo") = IIf(ListView1.ListItems.Item(ListView1.SelectedItem.Index).SmallIcon = 1, True, False)
            RsMail.Update
            
    
            imgMail = IIf(ListView1.ListItems.Item(ListView1.SelectedItem.Index).SmallIcon = 1, 2, 1)
            ListView1.ListItems.Item(ListView1.SelectedItem.Index).SmallIcon = imgMail
    End If
    RsMail.Close
End Sub
Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)

    'Verifica se SortKey é a mesma que a atual
    If ListView1.SortKey <> ColumnHeader.Index - 1 Then
            'Quando clicar em uma coluna define sortkey para indice -1
            ListView1.SortKey = ColumnHeader.Index - 1
            ListView1.SortOrder = lvwAscending
        Else
            'Se a coluna ja esta selecionada entao altera a
            'propr. SetOrder para ser o oposto da coluna em uso
            ListView1.SortOrder = IIf(ListView1.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    End If

    'Define a propriedade Sorted para utilizar a ordem atual
    ListView1.Sorted = True
End Sub

Private Sub hdMenu(op As Boolean)

    Tb_Menu.Buttons(1).Enabled = op
    'Tb_Menu.Buttons(2).Enabled = op
    Tb_Menu.Buttons(3).Enabled = op
    'Tb_Menu.Buttons(4).Enabled = op
    
    'Tb_Menu.Buttons(6).Enabled = IIf(op = False, True, False)
    'Tb_Menu.Buttons(7).Enabled = IIf(op = False, True, False)
    
End Sub

Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            'Acao = 1
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            Form_MailEnviar.Show
            'HDForm (True)
            'hdMenu (False)
            'LimpForm
        Case 2 'Alterar
            'Acao = 2
            'If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            'If Trim(txtDescr.Text) = "" Then Exit Sub
            'HDForm (True)
            'hdMenu (False)
        Case 3 'Excluir
            'Acao = 3
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            ExcluirMail
            
            'LimpForm
            'ExcluirCampo
            
        Case 4 'Imprimir
            'Acao = 4
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            'ImprimirListagem
        Case 6 'Gravar
            'GravarDados
            'LimpForm
            'HDForm (False)
            'hdMenu (True)
            'MstDados
        Case 7 'Cancelar
            'Acao = 7
            'HDForm (False)
            'hdMenu (True)
            'LimpForm
        Case 8 'Atualiza
            ExibirGrid
    End Select
End Sub

Private Sub ExibirMail(IDMail As Integer)
    Dim RsTMP As Recordset
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Mail WHERE ID = " & IDMail)
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Erro ao localizar Mensagem!", vbInformation, "CESNet - Aviso"
        Else
            RsTMP.MoveFirst
            lbData.Caption = RsTMP.Fields("Data")
            lbDe.Caption = PgRespUsu(RsTMP.Fields("De"))
            lbAssunto.Caption = RsTMP.Fields("Assunto")
            txtDescr.Text = RsTMP.Fields("Descr")
    End If
    RsTMP.Close
End Sub



Private Sub txtDescr_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub ExcluirMail()
    If MailID = 0 Then Exit Sub
    If MsgBox("Deseja realmente EXCLUIR esta Mensagem!", vbInformation + vbYesNo, "CESNet - Aviso") = vbYes Then
        BD.Execute "DELETE * FROM Mail WHERE ID = " & MailID
        MsgBox "Mensagem Excluida!", vbExclamation, "CESNet - Aviso"
        ExibirGrid
        LimpForm
    End If
End Sub
