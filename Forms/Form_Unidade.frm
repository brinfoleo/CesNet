VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Unidade 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Cadastro de Unidades"
   ClientHeight    =   8145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6840
   Icon            =   "Form_Unidade.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8145
   ScaleWidth      =   6840
   Begin VB.TextBox Txt_NomeCompleto 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1350
      MaxLength       =   100
      TabIndex        =   40
      Top             =   3180
      Width           =   5280
   End
   Begin VB.TextBox Txt_Nome 
      Enabled         =   0   'False
      Height          =   285
      Left            =   1365
      MaxLength       =   50
      TabIndex        =   19
      Top             =   2820
      Width           =   5280
   End
   Begin MSMask.MaskEdBox MEB_UnidID 
      Height          =   315
      Left            =   1365
      TabIndex        =   18
      Top             =   2460
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   556
      _Version        =   393216
      PromptInclude   =   0   'False
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   3
      Mask            =   "###"
      PromptChar      =   "_"
   End
   Begin VB.Frame Frame3 
      Height          =   2055
      Left            =   60
      TabIndex        =   5
      Top             =   3585
      Width           =   6720
      Begin VB.TextBox Txt_CxP 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3000
         MaxLength       =   20
         TabIndex        =   25
         Top             =   900
         Width           =   1515
      End
      Begin VB.TextBox Txt_Mail 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   26
         Top             =   1260
         Width           =   2475
      End
      Begin VB.ComboBox CB_UF 
         Enabled         =   0   'False
         Height          =   315
         ItemData        =   "Form_Unidade.frx":030A
         Left            =   5760
         List            =   "Form_Unidade.frx":035F
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   540
         Width           =   750
      End
      Begin VB.TextBox Txt_Mun 
         Enabled         =   0   'False
         Height          =   285
         Left            =   3660
         MaxLength       =   30
         TabIndex        =   22
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox Txt_Bai 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   30
         TabIndex        =   21
         Top             =   540
         Width           =   1695
      End
      Begin VB.TextBox Txt_End 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   50
         TabIndex        =   20
         Top             =   180
         Width           =   5595
      End
      Begin MSMask.MaskEdBox MEB_Cep 
         Height          =   255
         Left            =   900
         TabIndex        =   24
         Top             =   900
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
      Begin MSMask.MaskEdBox Meb_Tel1 
         Height          =   315
         Left            =   4920
         TabIndex        =   27
         Top             =   1260
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   14
         Mask            =   "(##) ####-####"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox Meb_Tel2 
         Height          =   315
         Left            =   4920
         TabIndex        =   28
         Top             =   1620
         Width           =   1575
         _ExtentX        =   2778
         _ExtentY        =   556
         _Version        =   393216
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   14
         Mask            =   "(##) ####-####"
         PromptChar      =   "_"
      End
      Begin VB.Label Label14 
         Alignment       =   1  'Right Justify
         Caption         =   "E-mail:"
         Height          =   195
         Left            =   300
         TabIndex        =   17
         Top             =   1320
         Width           =   495
      End
      Begin VB.Label Label9 
         Alignment       =   1  'Right Justify
         Caption         =   "Tel.(s):"
         Height          =   195
         Left            =   4320
         TabIndex        =   16
         Top             =   1500
         Width           =   495
      End
      Begin VB.Label Label8 
         Alignment       =   1  'Right Justify
         Caption         =   "Cx. Postal:"
         Height          =   195
         Left            =   2100
         TabIndex        =   15
         Top             =   960
         Width           =   795
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "CEP:"
         Height          =   195
         Left            =   420
         TabIndex        =   14
         Top             =   900
         Width           =   375
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "UF:"
         Height          =   195
         Left            =   5460
         TabIndex        =   13
         Top             =   600
         Width           =   255
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Municipio:"
         Height          =   195
         Left            =   2880
         TabIndex        =   12
         Top             =   600
         Width           =   735
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Bairro:"
         Height          =   195
         Left            =   360
         TabIndex        =   11
         Top             =   540
         Width           =   435
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Endereço:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   180
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Documentos:"
      Height          =   2355
      Left            =   90
      TabIndex        =   4
      Top             =   5655
      Width           =   6675
      Begin VB.TextBox Txt_CodEscolar 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   10
         TabIndex        =   42
         Top             =   990
         Width           =   1995
      End
      Begin VB.TextBox Txt_AutoCurso 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   38
         Top             =   1890
         Width           =   4275
      End
      Begin VB.TextBox Txt_Criacao 
         Enabled         =   0   'False
         Height          =   285
         Left            =   2280
         MaxLength       =   30
         TabIndex        =   37
         Top             =   1410
         Width           =   4275
      End
      Begin VB.TextBox Txt_UA 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1200
         MaxLength       =   20
         TabIndex        =   31
         Top             =   945
         Width           =   1995
      End
      Begin VB.TextBox Txt_IM 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   33
         Top             =   630
         Width           =   1995
      End
      Begin VB.TextBox Txt_IE 
         Enabled         =   0   'False
         Height          =   285
         Left            =   4560
         MaxLength       =   20
         TabIndex        =   32
         Top             =   225
         Width           =   1995
      End
      Begin VB.TextBox Txt_CNPJ 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   30
         Top             =   585
         Width           =   1995
      End
      Begin VB.TextBox Txt_CensoF 
         Enabled         =   0   'False
         Height          =   285
         Left            =   1215
         MaxLength       =   20
         TabIndex        =   29
         Top             =   240
         Width           =   1995
      End
      Begin VB.Label Label19 
         Caption         =   "Cód. Escolar:"
         Height          =   195
         Left            =   3495
         TabIndex        =   41
         Top             =   1035
         Width           =   960
      End
      Begin VB.Label Label17 
         Alignment       =   1  'Right Justify
         Caption         =   "Autorização do Curso(ato, número e data):"
         Height          =   495
         Left            =   180
         TabIndex        =   36
         Top             =   1770
         Width           =   1995
      End
      Begin VB.Label Label16 
         Alignment       =   1  'Right Justify
         Caption         =   "Ato de Criação:"
         Height          =   195
         Left            =   1020
         TabIndex        =   35
         Top             =   1470
         Width           =   1095
      End
      Begin VB.Label Label15 
         Alignment       =   1  'Right Justify
         Caption         =   "U. A.:"
         Height          =   195
         Left            =   675
         TabIndex        =   34
         Top             =   1035
         Width           =   420
      End
      Begin VB.Label Label13 
         Alignment       =   1  'Right Justify
         Caption         =   "Censo Federal:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   1095
      End
      Begin VB.Label Label12 
         Alignment       =   1  'Right Justify
         Caption         =   "Insc. Mun.:"
         Height          =   195
         Left            =   3660
         TabIndex        =   8
         Top             =   660
         Width           =   795
      End
      Begin VB.Label Label11 
         Alignment       =   1  'Right Justify
         Caption         =   "Insc. Est.:"
         Height          =   195
         Left            =   3720
         TabIndex        =   7
         Top             =   300
         Width           =   735
      End
      Begin VB.Label Label10 
         Alignment       =   1  'Right Justify
         Caption         =   "CNPJ:"
         Height          =   195
         Left            =   660
         TabIndex        =   6
         Top             =   660
         Width           =   435
      End
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   60
      TabIndex        =   0
      Top             =   960
      Width           =   6720
      Begin MSFlexGridLib.MSFlexGrid MSFG_Unidades 
         Height          =   1155
         Left            =   120
         TabIndex        =   1
         ToolTipText     =   "Para selecionar de um duplo click."
         Top             =   180
         Width           =   6495
         _ExtentX        =   11456
         _ExtentY        =   2037
         _Version        =   393216
         HighLight       =   0
         ScrollBars      =   2
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   "^Unidade ID    |^Unidades                                                                                          "
      End
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   43
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
         Left            =   5460
         Top             =   0
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
               Picture         =   "Form_Unidade.frx":03CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":06E8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":0A02
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":0D1C
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":1036
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":1350
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":166A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":1984
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":1C9E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":1FB8
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":22D2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":25EC
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":2906
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":2C20
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":2F3A
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":3254
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":356E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":3888
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":3BA2
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":3EBC
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":41D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":44F0
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":480A
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":4B24
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":4E3E
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":5158
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":5472
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":578C
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":5AA6
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":5DC0
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":60DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_Unidade.frx":63F4
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin VB.Label Label20 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CADASTRO DE UNIDADE"
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
      TabIndex        =   44
      Top             =   600
      Width           =   6855
   End
   Begin VB.Label Label18 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome Completo:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   90
      TabIndex        =   39
      Top             =   3225
      Width           =   1185
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Unidade:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   480
      TabIndex        =   3
      Top             =   2880
      Width           =   795
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Unidade ID:"
      ForeColor       =   &H00000000&
      Height          =   195
      Left            =   240
      TabIndex        =   2
      Top             =   2520
      Width           =   1035
   End
End
Attribute VB_Name = "Form_Unidade"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsUnid As Recordset
Dim lin, tmp As String
Dim YesNo As String
Private Sub GrvDados()
    If MEB_UnidID.Text = "" Or Txt_Nome.Text = "" Then
        MsgBox "Campos UNIDADE ID ou UNIDADE não devem ficar em Branco, por favor verifique!", vbExclamation, "CESNet - Aviso"
        MEB_UnidID.SetFocus
        Exit Sub
    End If
    With RsUnid
        Select Case tmp
            Case 1 'Incluir
                .AddNew
                .Fields("UnidID") = MEB_UnidID.Text
                .Fields("Nome") = Trim(Txt_Nome.Text)
                .Fields("NomeCompleto") = Trim(Txt_NomeCompleto.Text)
                .Fields("End") = IIf(Txt_End.Text = "", Null, Txt_End.Text)
                .Fields("Bai") = IIf(Txt_Bai.Text = "", Null, Txt_Bai.Text)
                .Fields("Mun") = IIf(Txt_Mun.Text = "", Null, Txt_Mun.Text)
                .Fields("UF") = IIf(CB_UF.Text = "", Null, CB_UF.Text)
                .Fields("CEP") = IIf(MEB_Cep.Text = "", Null, MEB_Cep.Text)
                .Fields("CxP") = IIf(Txt_CxP.Text = "", Null, Txt_CxP.Text)
                .Fields("Tel1") = IIf(Meb_Tel1.Text = "", Null, Meb_Tel1.Text)
                .Fields("Tel2") = IIf(Meb_Tel2.Text = "", Null, Meb_Tel2.Text)
                .Fields("Mail") = IIf(Txt_Mail.Text = "", Null, Txt_Mail.Text)
                .Fields("CF") = IIf(Txt_CensoF.Text = "", Null, Txt_CensoF.Text)
                .Fields("CNPJ") = IIf(Txt_CNPJ.Text = "", Null, Txt_CNPJ.Text)
                .Fields("IE") = IIf(Txt_IE.Text = "", Null, Txt_IE.Text)
                .Fields("IM") = IIf(Txt_IM.Text = "", Null, Txt_IM.Text)
                .Fields("UA") = IIf(Txt_UA.Text = "", Null, Txt_UA.Text)
                .Fields("AutoCurso") = IIf(Trim(Txt_AutoCurso.Text) = "", Null, Trim(Txt_AutoCurso.Text))
                .Fields("Criacao") = IIf(Trim(Txt_Criacao.Text) = "", Null, Trim(Txt_Criacao.Text))
                .Fields("CodEscolar") = IIf(Trim(Txt_CodEscolar.Text) = "", Null, Trim(Txt_CodEscolar.Text))
                .Update
                LstUnidades
                
            
            Case 2 'Alterar
                .Edit
                .Fields("UnidID") = MEB_UnidID.Text
                .Fields("Nome") = Trim(Txt_Nome.Text)
                .Fields("NomeCompleto") = Trim(Txt_NomeCompleto.Text)
                .Fields("End") = IIf(Txt_End.Text = "", Null, Txt_End.Text)
                .Fields("Bai") = IIf(Txt_Bai.Text = "", Null, Txt_Bai.Text)
                .Fields("Mun") = IIf(Txt_Mun.Text = "", Null, Txt_Mun.Text)
                .Fields("UF") = IIf(CB_UF.Text = "", Null, CB_UF.Text)
                .Fields("CEP") = IIf(MEB_Cep.Text = "", Null, MEB_Cep.Text)
                .Fields("CxP") = IIf(Txt_CxP.Text = "", Null, Txt_CxP.Text)
                .Fields("Tel1") = IIf(Meb_Tel1.Text = "", Null, Meb_Tel1.Text)
                .Fields("Tel2") = IIf(Meb_Tel2.Text = "", Null, Meb_Tel2.Text)
                .Fields("Mail") = IIf(Txt_Mail.Text = "", Null, Txt_Mail.Text)
                .Fields("CF") = IIf(Txt_CensoF.Text = "", Null, Txt_CensoF.Text)
                .Fields("CNPJ") = IIf(Txt_CNPJ.Text = "", Null, Txt_CNPJ.Text)
                .Fields("IE") = IIf(Txt_IE.Text = "", Null, Txt_IE.Text)
                .Fields("IM") = IIf(Txt_IM.Text = "", Null, Txt_IM.Text)
                .Fields("UA") = IIf(Txt_UA.Text = "", Null, Txt_UA.Text)
                .Fields("AutoCurso") = IIf(Trim(Txt_AutoCurso.Text) = "", Null, Trim(Txt_AutoCurso.Text))
                .Fields("Criacao") = IIf(Trim(Txt_Criacao.Text) = "", Null, Trim(Txt_Criacao.Text))
                .Fields("CodEscolar") = IIf(Trim(Txt_CodEscolar.Text) = "", Null, Trim(Txt_CodEscolar.Text))
                .Update
                LstUnidades
        End Select
    End With
    tmp = 0
   
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
    tmp = 0
    hdMenu (True)
    Set RsUnid = BD.OpenRecordset("SELECT * FROM Unidades ORDER BY UnidID")
    If RsUnid.BOF And RsUnid.EOF Then
        Else
            LstUnidades
            RsUnid.MoveFirst
            MstUnidade
    End If
End Sub

Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            HDForm (True)
            hdMenu (False)
            LimpForm
            MSFG_Unidades.Enabled = False
            tmp = 1
            MEB_UnidID.SetFocus
        Case 2 'Alterar
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            If MEB_UnidID.Text = "" Then
                MsgBox "É necessário que selecione uma Unidade.", vbInformation, "CESNet - Aviso!"
                Exit Sub
            End If
            MSFG_Unidades.Enabled = False
            tmp = 2
            'If Trim(txtDescr.Text) = "" Then Exit Sub
            HDForm (True)
            hdMenu (False)
        Case 3 'Excluir
            tmp = 3
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            If MEB_UnidID.Text = "" Then
                MsgBox "É necessário que selecione uma Unidade.", vbInformation, "CESNet - Aviso!"
                Exit Sub
            End If
            With RsUnid
                .FindFirst "UnidID ='" & MEB_UnidID.Text & "'"
                If .NoMatch Then
                        MsgBox "Erro no Acesso aos dados. Por favor abra o formulário novamente.", vbExclamation, "CESNet - Aviso"
                        Unload Me
                    Else
                        YesNo = MsgBox("Deseja EXCLUIR o Registro nº " & MEB_UnidID.Text, vbYesNo, "CESNet - Aviso")
                        If YesNo = 6 Then
                            .Delete
                            LstUnidades
                            LimpForm
                        End If
                End If
            End With

        
        Case 4 'Imprimir
            tmp = 4
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            'ImprimirListagem
        Case 6 'Gravar
            If ValidarSoftware("Unidades") = False Then Exit Sub
            GrvDados
            LimpForm
            HDForm (False)
            hdMenu (True)
            MSFG_Unidades.Enabled = True
            LstUnidades
        Case 7 'Cancelar
            MSFG_Unidades.Enabled = True
            tmp = 0
            HDForm (False)
            hdMenu (True)
            LimpForm
    End Select
End Sub
Private Sub hdMenu(op As Boolean)

    Tb_Menu.Buttons(1).Enabled = op
    Tb_Menu.Buttons(2).Enabled = op
    Tb_Menu.Buttons(3).Enabled = op
    Tb_Menu.Buttons(4).Enabled = False
    
    Tb_Menu.Buttons(6).Enabled = IIf(op = False, True, False)
    Tb_Menu.Buttons(7).Enabled = IIf(op = False, True, False)
    
End Sub


Private Sub LimpForm()
    MEB_UnidID.Text = ""
    Txt_Nome.Text = ""
    Txt_NomeCompleto.Text = ""
    Txt_End.Text = ""
    Txt_Bai.Text = ""
    Txt_Mun.Text = ""
    'CB_UF.Text = ""
    MEB_Cep.PromptInclude = False
    Meb_Tel1.PromptInclude = False
    Meb_Tel2.PromptInclude = False
    MEB_Cep.Text = ""
    Txt_CxP.Text = ""
    Txt_Mail.Text = ""
    Meb_Tel1.Text = ""
    Meb_Tel2.Text = ""
    
    Txt_CensoF.Text = ""
    Txt_CNPJ.Text = ""
    Txt_IE.Text = ""
    Txt_IM.Text = ""
    Txt_UA.Text = ""
    Txt_Criacao.Text = ""
    Txt_AutoCurso.Text = ""
    Txt_CodEscolar.Text = ""
    
    MEB_Cep.PromptInclude = True
    Meb_Tel1.PromptInclude = True
    Meb_Tel2.PromptInclude = True
    
End Sub
Private Sub HDForm(op As Boolean)
    MEB_UnidID.Enabled = op
    Txt_Nome.Enabled = op
    Txt_NomeCompleto.Enabled = op
    Txt_End.Enabled = op
    Txt_Bai.Enabled = op
    Txt_Mun.Enabled = op
    CB_UF.Enabled = op
    MEB_Cep.Enabled = op
    Txt_CxP.Enabled = op
    Txt_Mail.Enabled = op
    Meb_Tel1.Enabled = op
    Meb_Tel2.Enabled = op
    
    Txt_CensoF.Enabled = op
    Txt_CNPJ.Enabled = op
    Txt_IE.Enabled = op
    Txt_IM.Enabled = op
    Txt_UA.Enabled = op
    Txt_AutoCurso.Enabled = op
    Txt_Criacao.Enabled = op
    Txt_CodEscolar.Enabled = op
End Sub
Private Sub Form_Unload(Cancel As Integer)
    RsUnid.Close
End Sub


Private Sub MEB_UnidID_LostFocus()
    MEB_UnidID.Text = Mid(String(3, "0"), 1, 3 - Len(MEB_UnidID.Text)) & MEB_UnidID.Text
End Sub

Private Sub MSFG_Unidades_DblClick()
    With MSFG_Unidades
        If .TextMatrix(.Row, 0) = "" Or .TextMatrix(.Row, 0) = "Unidade ID" Then
            Exit Sub
        End If
        RsUnid.FindFirst "UnidID = '" & .TextMatrix(.Row, 0) & "'"
        If RsUnid.NoMatch Then
                MsgBox "Erro no acesso por favor reabra o formulário", vbExclamation, "CESNet - Aviso"
                Unload Me
            Else
                MstUnidade
        End If
    End With
End Sub

Private Sub Txt_Nome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Sub LstUnidades()
    lin = 1
    RsUnid.MoveFirst
    With MSFG_Unidades
        .Rows = 1
        While RsUnid.EOF = False
            .Rows = .Rows + 1
            .TextMatrix(lin, 0) = RsUnid.Fields("UnidID")
            .TextMatrix(lin, 1) = RsUnid.Fields("Nome")
            lin = lin + 1
            RsUnid.MoveNext
        Wend
        .ColSel = 0
        .Sort = 5
    End With
End Sub
Private Sub MstUnidade()
    With RsUnid
        MEB_UnidID.Text = .Fields("UnidID")
        Txt_Nome.Text = .Fields("Nome")
        Txt_NomeCompleto.Text = IIf(IsNull(.Fields("NomeCompleto")), "", .Fields("NomeCompleto"))
        Txt_End.Text = IIf(IsNull(.Fields("End")), "", .Fields("End"))
        Txt_Bai.Text = IIf(IsNull(.Fields("Bai")), "", .Fields("Bai"))
        Txt_Mun.Text = IIf(IsNull(.Fields("Mun")), "", .Fields("Mun"))
        CB_UF.Text = IIf(IsNull(.Fields("UF")), " ", .Fields("UF"))
        MEB_Cep.PromptInclude = False
        MEB_Cep.Text = Format(.Fields("CEP"), "#####-###")
        MEB_Cep.PromptInclude = True
        Txt_CxP.Text = IIf(IsNull(.Fields("Cxp")), "", .Fields("Cxp"))
        Meb_Tel1.PromptInclude = False
        Meb_Tel2.PromptInclude = False
        Meb_Tel1.Text = .Fields("Tel1")
        Meb_Tel2.Text = .Fields("Tel2")
        Meb_Tel1.PromptInclude = True
        Meb_Tel2.PromptInclude = True
        Txt_Mail.Text = IIf(IsNull(.Fields("Mail")), "", .Fields("Mail"))
        Txt_CensoF.Text = IIf(IsNull(.Fields("CF")), "", .Fields("CF"))
        Txt_CNPJ.Text = IIf(IsNull(.Fields("CNPJ")), "", .Fields("CNPJ"))
        Txt_IE.Text = IIf(IsNull(.Fields("IE")), "", .Fields("IE"))
        Txt_IM.Text = IIf(IsNull(.Fields("IM")), "", .Fields("IM"))
        Txt_UA.Text = IIf(IsNull(.Fields("UA")), "", .Fields("UA"))
        Txt_Criacao.Text = IIf(IsNull(.Fields("Criacao")), "", .Fields("Criacao"))
        Txt_AutoCurso.Text = IIf(IsNull(.Fields("AutoCurso")), "", .Fields("AutoCurso"))
        Txt_CodEscolar.Text = IIf(IsNull(.Fields("CodEscolar")), "", .Fields("CodEscolar"))
    End With
End Sub
Private Sub Meb_UnidID_GotFocus()
    MEB_UnidID.SelStart = 0
    MEB_UnidID.SelLength = 3
End Sub
Private Sub Txt_Nome_GotFocus()
    Txt_Nome.SelStart = 0
    Txt_Nome.SelLength = Len(Txt_Nome.Text)
End Sub
