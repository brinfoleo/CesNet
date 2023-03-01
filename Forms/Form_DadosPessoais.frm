VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_DadosPessoais 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Dados Pessoais"
   ClientHeight    =   8265
   ClientLeft      =   450
   ClientTop       =   465
   ClientWidth     =   10485
   Icon            =   "Form_DadosPessoais.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   8265
   ScaleWidth      =   10485
   Begin TabDlg.SSTab SST_Geral 
      Height          =   6255
      Left            =   45
      TabIndex        =   4
      Top             =   1845
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   11033
      _Version        =   393216
      TabOrientation  =   1
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Dados Pessoais"
      TabPicture(0)   =   "Form_DadosPessoais.frx":030A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame_DP"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Observações"
      TabPicture(1)   =   "Form_DadosPessoais.frx":0326
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame6"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Historico"
      TabPicture(2)   =   "Form_DadosPessoais.frx":0342
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame11"
      Tab(2).ControlCount=   1
      Begin VB.Frame Frame_DP 
         Height          =   5715
         Left            =   120
         TabIndex        =   28
         Top             =   60
         Width           =   10095
         Begin VB.Frame Frame3 
            Caption         =   "Endereço:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   2145
            Left            =   120
            TabIndex        =   48
            Top             =   480
            Width           =   9795
            Begin VB.ComboBox CB_UF 
               Enabled         =   0   'False
               Height          =   315
               ItemData        =   "Form_DadosPessoais.frx":035E
               Left            =   990
               List            =   "Form_DadosPessoais.frx":03B3
               Sorted          =   -1  'True
               Style           =   2  'Dropdown List
               TabIndex        =   57
               Top             =   1695
               Width           =   750
            End
            Begin VB.TextBox Txt_Mun 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4185
               MaxLength       =   30
               TabIndex        =   56
               Top             =   1335
               Width           =   3075
            End
            Begin VB.TextBox Txt_Bai 
               Enabled         =   0   'False
               Height          =   285
               Left            =   990
               MaxLength       =   30
               TabIndex        =   55
               Top             =   1335
               Width           =   1995
            End
            Begin VB.TextBox Txt_End 
               Enabled         =   0   'False
               Height          =   285
               Left            =   990
               MaxLength       =   50
               TabIndex        =   52
               Top             =   585
               Width           =   6270
            End
            Begin VB.TextBox Txt_Num 
               Enabled         =   0   'False
               Height          =   285
               Left            =   990
               MaxLength       =   30
               TabIndex        =   53
               Top             =   945
               Width           =   1140
            End
            Begin VB.TextBox Txt_Comp 
               Enabled         =   0   'False
               Height          =   285
               Left            =   4185
               MaxLength       =   30
               TabIndex        =   54
               Top             =   945
               Width           =   1635
            End
            Begin MSMask.MaskEdBox Meb_Cep 
               Height          =   255
               Left            =   990
               TabIndex        =   51
               Top             =   270
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
            Begin VB.Label Label7 
               Alignment       =   1  'Right Justify
               Caption         =   "CEP:"
               Height          =   195
               Left            =   570
               TabIndex        =   70
               Top             =   270
               Width           =   375
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "UF:"
               Height          =   195
               Left            =   645
               TabIndex        =   68
               Top             =   1755
               Width           =   255
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Municipio:"
               Height          =   195
               Left            =   3345
               TabIndex        =   66
               Top             =   1350
               Width           =   735
            End
            Begin VB.Label Label8 
               Alignment       =   1  'Right Justify
               Caption         =   "Bairro:"
               Height          =   195
               Left            =   465
               TabIndex        =   64
               Top             =   1350
               Width           =   435
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "Endereço:"
               Height          =   195
               Left            =   165
               TabIndex        =   62
               Top             =   570
               Width           =   735
            End
            Begin VB.Label Label50 
               Alignment       =   1  'Right Justify
               Caption         =   "Número:"
               Height          =   195
               Left            =   315
               TabIndex        =   60
               Top             =   1005
               Width           =   600
            End
            Begin VB.Label Label51 
               Alignment       =   1  'Right Justify
               Caption         =   "Complemento:"
               Height          =   195
               Left            =   3015
               TabIndex        =   58
               Top             =   990
               Width           =   1095
            End
         End
         Begin TabDlg.SSTab SST_DadosPessoais 
            Height          =   2805
            Left            =   120
            TabIndex        =   59
            Top             =   2760
            Width           =   9795
            _ExtentX        =   17277
            _ExtentY        =   4948
            _Version        =   393216
            TabHeight       =   520
            TabCaption(0)   =   "Documentos"
            TabPicture(0)   =   "Form_DadosPessoais.frx":0422
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame_Proprio"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Contatos"
            TabPicture(1)   =   "Form_DadosPessoais.frx":043E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame1"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Outros"
            TabPicture(2)   =   "Form_DadosPessoais.frx":045A
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame22"
            Tab(2).Control(1)=   "Frame2"
            Tab(2).ControlCount=   2
            Begin VB.Frame Frame2 
               Caption         =   "Filiação:"
               Height          =   1035
               Left            =   -74820
               TabIndex        =   45
               Top             =   405
               Width           =   9315
               Begin VB.TextBox Txt_Mae 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   600
                  MaxLength       =   50
                  TabIndex        =   79
                  Top             =   240
                  Width           =   6465
               End
               Begin VB.TextBox Txt_Pai 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   600
                  MaxLength       =   50
                  TabIndex        =   80
                  Top             =   600
                  Width           =   6465
               End
               Begin MSMask.MaskEdBox mebDtNascMae 
                  Height          =   315
                  Left            =   7860
                  TabIndex        =   91
                  Top             =   240
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  Enabled         =   0   'False
                  MaxLength       =   10
                  Format          =   "DD/MM/YYYY"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox mebDtNascPai 
                  Height          =   315
                  Left            =   7860
                  TabIndex        =   92
                  Top             =   600
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  Enabled         =   0   'False
                  MaxLength       =   10
                  Format          =   "DD/MM/YYYY"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Nasc.:"
                  Height          =   195
                  Left            =   7200
                  TabIndex        =   90
                  Top             =   660
                  Width           =   495
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Nasc.:"
                  Height          =   195
                  Left            =   7200
                  TabIndex        =   89
                  Top             =   300
                  Width           =   495
               End
               Begin VB.Label Label27 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Mãe:"
                  Height          =   255
                  Left            =   180
                  TabIndex        =   47
                  Top             =   300
                  Width           =   375
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Pai:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   46
                  Top             =   660
                  Width           =   315
               End
            End
            Begin VB.Frame Frame_Proprio 
               Caption         =   "Documentos:"
               Height          =   2310
               Left            =   90
               TabIndex        =   36
               Top             =   315
               Width           =   9435
               Begin VB.ComboBox CB_EstCiv 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_DadosPessoais.frx":0476
                  Left            =   5835
                  List            =   "Form_DadosPessoais.frx":0492
                  Style           =   2  'Dropdown List
                  TabIndex        =   69
                  Top             =   195
                  Width           =   3180
               End
               Begin VB.TextBox Txt_Nat 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5835
                  MaxLength       =   30
                  TabIndex        =   71
                  Top             =   735
                  Width           =   3180
               End
               Begin VB.TextBox Txt_OrgEmi 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   20
                  TabIndex        =   63
                  Top             =   750
                  Width           =   2910
               End
               Begin VB.TextBox Txt_RG 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   20
                  TabIndex        =   61
                  Top             =   240
                  Width           =   2895
               End
               Begin VB.TextBox Txt_CertNasc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   20
                  TabIndex        =   67
                  Top             =   1800
                  Width           =   2910
               End
               Begin VB.TextBox Txt_Nac 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5805
                  MaxLength       =   30
                  TabIndex        =   74
                  Top             =   1770
                  Width           =   3180
               End
               Begin VB.TextBox Txt_CPF 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1200
                  MaxLength       =   20
                  TabIndex        =   65
                  Top             =   1230
                  Width           =   2910
               End
               Begin VB.ComboBox CB_UFNatural 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_DadosPessoais.frx":04ED
                  Left            =   5820
                  List            =   "Form_DadosPessoais.frx":0542
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   73
                  Top             =   1200
                  Width           =   750
               End
               Begin VB.Label Label44 
                  Alignment       =   1  'Right Justify
                  Caption         =   "CPF:"
                  Height          =   195
                  Left            =   825
                  TabIndex        =   44
                  Top             =   1275
                  Width           =   330
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Natural da Cidade:"
                  Height          =   375
                  Left            =   4830
                  TabIndex        =   43
                  Top             =   615
                  Width           =   960
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Estado Civil:"
                  Height          =   195
                  Left            =   4875
                  TabIndex        =   42
                  Top             =   255
                  Width           =   915
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Orgão Emissor:"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   41
                  Top             =   810
                  Width           =   1065
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Caption         =   "RG:"
                  Height          =   195
                  Left            =   855
                  TabIndex        =   40
                  Top             =   300
                  Width           =   300
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Certidão de Nascimento:"
                  Height          =   375
                  Left            =   285
                  TabIndex        =   39
                  Top             =   1665
                  Width           =   870
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Nacionalidade:"
                  Height          =   195
                  Left            =   4665
                  TabIndex        =   38
                  Top             =   1800
                  Width           =   1095
               End
               Begin VB.Label Label52 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Natural da UF:"
                  Height          =   420
                  Left            =   4875
                  TabIndex        =   37
                  Top             =   1140
                  Width           =   870
               End
            End
            Begin VB.Frame Frame1 
               Caption         =   "Contatos:"
               Height          =   2280
               Left            =   -74850
               TabIndex        =   32
               Top             =   390
               Width           =   9345
               Begin VB.TextBox Txt_Mail 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   765
                  MaxLength       =   30
                  TabIndex        =   75
                  Top             =   360
                  Width           =   3375
               End
               Begin VB.TextBox Txt_Tel1 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   765
                  MaxLength       =   14
                  TabIndex        =   77
                  Top             =   1170
                  Width           =   1875
               End
               Begin VB.TextBox Txt_Cel 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   765
                  MaxLength       =   14
                  TabIndex        =   76
                  Top             =   750
                  Width           =   1905
               End
               Begin VB.TextBox Txt_Tel2 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   765
                  MaxLength       =   14
                  TabIndex        =   78
                  Top             =   1560
                  Width           =   1875
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Caption         =   "E-mail:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   35
                  Top             =   405
                  Width           =   450
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tel.(s):"
                  Height          =   195
                  Left            =   135
                  TabIndex        =   34
                  Top             =   1350
                  Width           =   495
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Celular:"
                  Height          =   255
                  Left            =   75
                  TabIndex        =   33
                  Top             =   780
                  Width           =   555
               End
            End
            Begin VB.Frame Frame22 
               Caption         =   "Outros:"
               Height          =   1155
               Left            =   -74820
               TabIndex        =   29
               Top             =   1530
               Width           =   9450
               Begin VB.TextBox txtTpSang 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   8520
                  MaxLength       =   5
                  TabIndex        =   94
                  Text            =   "Text1"
                  Top             =   300
                  Width           =   615
               End
               Begin VB.TextBox txtRaca 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5580
                  MaxLength       =   60
                  TabIndex        =   88
                  Text            =   "Text1"
                  Top             =   240
                  Width           =   1575
               End
               Begin VB.TextBox txtOpcaoRel 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5580
                  MaxLength       =   60
                  TabIndex        =   87
                  Text            =   "Text1"
                  Top             =   600
                  Width           =   1575
               End
               Begin VB.ComboBox Cb_Deficiencia 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   1005
                  Style           =   2  'Dropdown List
                  TabIndex        =   83
                  Top             =   660
                  Width           =   3030
               End
               Begin VB.ComboBox Cb_Sexo 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_DadosPessoais.frx":05B1
                  Left            =   2940
                  List            =   "Form_DadosPessoais.frx":05BE
                  Style           =   2  'Dropdown List
                  TabIndex        =   82
                  Top             =   240
                  Width           =   1035
               End
               Begin MSMask.MaskEdBox Meb_Nasc 
                  Height          =   315
                  Left            =   1035
                  TabIndex        =   81
                  Top             =   225
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   556
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  Enabled         =   0   'False
                  MaxLength       =   10
                  Format          =   "DD/MM/YYYY"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tp.Sanguineo:"
                  Height          =   255
                  Left            =   7320
                  TabIndex        =   93
                  Top             =   300
                  Width           =   1035
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Opção Religiosa:"
                  Height          =   315
                  Left            =   4200
                  TabIndex        =   86
                  Top             =   660
                  Width           =   1275
               End
               Begin VB.Label Label23 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Raça:"
                  Height          =   195
                  Left            =   4500
                  TabIndex        =   85
                  Top             =   300
                  Width           =   975
               End
               Begin VB.Label Label41 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Deficiencia:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   84
                  Top             =   720
                  Width           =   870
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Nascimento:"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   31
                  Top             =   270
                  Width           =   915
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Sexo:"
                  Height          =   195
                  Left            =   2400
                  TabIndex        =   30
                  Top             =   300
                  Width           =   495
               End
            End
         End
         Begin VB.Label Label22 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "DADOS PESSOAIS"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   72
            Top             =   120
            Width           =   9975
         End
      End
      Begin VB.Frame Frame11 
         Height          =   5355
         Left            =   -74880
         TabIndex        =   9
         Top             =   60
         Width           =   10095
         Begin VB.ComboBox Cb_SerieHistEsc 
            Enabled         =   0   'False
            Height          =   315
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   3600
            Width           =   1695
         End
         Begin VB.TextBox Txt_EstabHistEsc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            MaxLength       =   50
            TabIndex        =   14
            Top             =   4440
            Width           =   6135
         End
         Begin VB.TextBox Txt_CidadeHistEsc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1440
            MaxLength       =   40
            TabIndex        =   13
            Top             =   4860
            Width           =   6135
         End
         Begin VB.TextBox Txt_AnoHistEsc 
            Enabled         =   0   'False
            Height          =   285
            Left            =   1410
            MaxLength       =   4
            TabIndex        =   12
            Top             =   4020
            Width           =   1035
         End
         Begin VB.CommandButton Bt_GravarHistEsc 
            Caption         =   "Gravar"
            Enabled         =   0   'False
            Height          =   810
            Left            =   7875
            Picture         =   "Form_DadosPessoais.frx":05CB
            Style           =   1  'Graphical
            TabIndex        =   11
            Top             =   3600
            Width           =   2085
         End
         Begin VB.CommandButton Bt_CancelarHistEsc 
            Caption         =   "Cancelar"
            Enabled         =   0   'False
            Height          =   810
            Left            =   7875
            Picture         =   "Form_DadosPessoais.frx":08D5
            Style           =   1  'Graphical
            TabIndex        =   10
            Top             =   4440
            Width           =   2085
         End
         Begin MSFlexGridLib.MSFlexGrid MSFG_HistEscolar 
            Height          =   2835
            Left            =   90
            TabIndex        =   16
            ToolTipText     =   "Pressione <Insert> para inserir linha ou <Delete> para apagar linha."
            Top             =   540
            Width           =   9720
            _ExtentX        =   17145
            _ExtentY        =   5001
            _Version        =   393216
            Cols            =   4
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"Form_DadosPessoais.frx":0BDF
         End
         Begin VB.Label Label32 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "HISTORICO ESCOLAR"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   22
            Top             =   120
            Width           =   9975
         End
         Begin VB.Label Label34 
            Alignment       =   1  'Right Justify
            Caption         =   "Série:"
            Height          =   195
            Left            =   900
            TabIndex        =   21
            Top             =   3660
            Width           =   435
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Ano:"
            Height          =   195
            Left            =   960
            TabIndex        =   20
            Top             =   4080
            Width           =   375
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Estabelecimento:"
            Height          =   195
            Left            =   120
            TabIndex        =   19
            Top             =   4500
            Width           =   1215
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "Cidade/Estado:"
            Height          =   195
            Left            =   180
            TabIndex        =   18
            Top             =   4920
            Width           =   1155
         End
         Begin VB.Label Label38 
            Caption         =   "Duplo click p/editar linha ou pressione <Insert> para inserir linha ou <Delete> para apagar linha."
            ForeColor       =   &H00404040&
            Height          =   195
            Left            =   3120
            TabIndex        =   17
            Top             =   3360
            Width           =   6795
         End
      End
      Begin VB.Frame Frame6 
         Height          =   5535
         Left            =   -74880
         TabIndex        =   7
         Top             =   120
         Width           =   10080
         Begin VB.CommandButton Bt_AltObs 
            Caption         =   "Alterar"
            Height          =   795
            Left            =   5640
            Picture         =   "Form_DadosPessoais.frx":0C8A
            Style           =   1  'Graphical
            TabIndex        =   24
            Top             =   4620
            Width           =   2115
         End
         Begin VB.CommandButton Bt_GrvObs 
            Caption         =   "Gravar"
            Enabled         =   0   'False
            Height          =   795
            Left            =   7815
            Picture         =   "Form_DadosPessoais.frx":0F94
            Style           =   1  'Graphical
            TabIndex        =   23
            Top             =   4620
            Width           =   2115
         End
         Begin VB.TextBox Txt_Obs 
            Enabled         =   0   'False
            Height          =   3930
            Left            =   90
            MaxLength       =   2000
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   8
            Top             =   540
            Width           =   9840
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "OBSERVAÇÕES"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H8000000E&
            Height          =   255
            Left            =   60
            TabIndex        =   25
            Top             =   120
            Width           =   9975
         End
      End
   End
   Begin VB.ComboBox CB_Unidade 
      Enabled         =   0   'False
      Height          =   315
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   50
      Top             =   1410
      Width           =   5370
   End
   Begin VB.ComboBox Cb_Nome 
      Height          =   315
      ItemData        =   "Form_DadosPessoais.frx":129E
      Left            =   765
      List            =   "Form_DadosPessoais.frx":12A0
      Sorted          =   -1  'True
      TabIndex        =   49
      Top             =   990
      Width           =   6675
   End
   Begin MSMask.MaskEdBox MebMatricula 
      Height          =   375
      Left            =   8730
      TabIndex        =   3
      Top             =   945
      Width           =   1635
      _ExtentX        =   2884
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##.###.####"
      PromptChar      =   "_"
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1058
      ButtonWidth     =   1323
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "IL_Menu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   13
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
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   2
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Ficha de Matricula"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Text            =   "Carteira"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Matricula"
            Key             =   "Matricula"
            ImageIndex      =   31
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Foto"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Capturar"
            ImageIndex      =   33
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList IL_Menu 
         Left            =   7320
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
               Picture         =   "Form_DadosPessoais.frx":12A2
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":15BC
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":18D6
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":1BF0
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":1F0A
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":2224
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":253E
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":2858
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":2B72
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":2E8C
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":31A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":34C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":37DA
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":3AF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":3E0E
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":4128
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":4442
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":475C
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":4A76
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":4D90
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":50AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":53C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":56DE
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":59F8
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":5D12
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":602C
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":6346
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":6660
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":697A
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":6C94
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":6FAE
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":72C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_DadosPessoais.frx":75E2
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSMask.MaskEdBox Meb_Dt 
      Height          =   315
      Left            =   9075
      TabIndex        =   26
      Top             =   1380
      Width           =   1275
      _ExtentX        =   2249
      _ExtentY        =   556
      _Version        =   393216
      AutoTab         =   -1  'True
      Enabled         =   0   'False
      MaxLength       =   10
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Format          =   "DD/MM/YYYY"
      Mask            =   "##/##/####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Dt. Cadastro:"
      Height          =   195
      Left            =   8100
      TabIndex        =   27
      Top             =   1455
      Width           =   930
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "DADOS PESSOAIS"
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
      Width           =   10515
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Unidade:"
      Height          =   195
      Left            =   60
      TabIndex        =   2
      Top             =   1515
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Matricula:"
      Height          =   195
      Left            =   7920
      TabIndex        =   1
      Top             =   1095
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   1050
      Width           =   495
   End
End
Attribute VB_Name = "Form_DadosPessoais"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim RsMatricula                 As Recordset
Dim RsUnidade                   As Recordset
Dim RsHistEscolar               As Recordset
Dim RsMatriculaEnsino           As Recordset
Dim RsMatriculaDisciplina       As Recordset
Dim RsMatriculaSerie            As Recordset
Dim RsMatriculaProva            As Recordset
Dim RsGradeEnsinoDisciplinas    As Recordset
Dim RsGradeEnsinoSerie          As Recordset
Dim RsTrafego                   As Recordset
Dim RsProvas                    As Recordset
Dim RsEnsino                    As Recordset
Dim RsDisciplina                As Recordset
Dim RsSerie                     As Recordset
Dim RsDeficiencia               As Recordset
'Dim RsCoordImpCert              As Recordset
Dim RsInstEnsino                As Recordset
Dim RsTMP                       As Recordset

'Dim Evento          As Boolean 'Inf. ao sis .LstSeries_Click se o siste esta listando ou se o usu esta clicando
Dim yn              As Integer  'usar para solicitar YES ou NO
Dim lin             As Integer ' usado para contar linhas nos MSFG´s

Dim MatrID          As String

Dim EnsinoID        As Integer
Dim DisciplinaID    As Integer
Dim SerieID         As Integer
Dim Ensino          As String
Dim Disciplina      As String
Dim Serie           As String
Dim RefTrafegoID As String 'referencia na Trafego


Dim Acao As String 'define se o usuario esta incluindo. alterando ou fazendo somente uma consulta
Dim tmp, cont As String
Dim tmp1, cont1 As String
'Variaveis para impressao de certificado
Dim Topo(20) As Integer
Dim MargE(20) As Integer







Private Sub Bt_CancelarHistEsc_Click()
    Cb_SerieHistEsc.Enabled = False
    Txt_AnoHistEsc.Enabled = False
    Txt_EstabHistEsc.Enabled = False
    Txt_CidadeHistEsc.Enabled = False
    
    MSFG_HistEscolar.Enabled = True

    Bt_GravarHistEsc.Enabled = False
    Bt_CancelarHistEsc.Enabled = False
    'Bt_ImprCert.Enabled = True
    For lin = MSFG_HistEscolar.Rows - 1 To 1 Step -1
        With MSFG_HistEscolar
            If .TextMatrix(lin, 0) = "" Then
                If .Rows = 2 Then
                        .Rows = 1
                        .Rows = 2
                    Else
                        .RemoveItem lin
                End If
            End If
        End With
    Next
End Sub

Private Sub Bt_GravarHistEsc_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub

    
    If Trim(Cb_SerieHistEsc.Text) = "" Then
        MsgBox "O campo SÉRIE não pode ser deixado em branco. ", vbInformation, "CESNet - Atenção"
        Exit Sub
    End If
    With MSFG_HistEscolar
        .TextMatrix(.Row, 0) = Cb_SerieHistEsc.Text
        .TextMatrix(.Row, 1) = Txt_AnoHistEsc.Text
        .TextMatrix(.Row, 2) = Txt_EstabHistEsc.Text
        .TextMatrix(.Row, 3) = Txt_CidadeHistEsc.Text
    End With
    Bt_CancelarHistEsc_Click
    BD.Execute "DELETE * FROM HistEscolar WHERE MatrID = '" & MatrID & "'"
    Set RsHistEscolar = BD.OpenRecordset("SELECT * FROM HistEscolar ORDER BY MatrID")
    With MSFG_HistEscolar
        For lin = 1 To .Rows - 1
            RsHistEscolar.AddNew
            RsHistEscolar.Fields("MatrID") = MatrID
            RsHistEscolar.Fields("Seq") = lin
            RsHistEscolar.Fields("Serie") = Trim(.TextMatrix(lin, 0))
            RsHistEscolar.Fields("Ano") = IIf(Trim(.TextMatrix(lin, 1)) = "", Null, Trim(.TextMatrix(lin, 1)))
            RsHistEscolar.Fields("Escola") = IIf(Trim(.TextMatrix(lin, 2)) = "", Null, Trim(.TextMatrix(lin, 2)))
            RsHistEscolar.Fields("Cidade") = IIf(Trim(.TextMatrix(lin, 3)) = "", Null, Trim(.TextMatrix(lin, 3)))
            RsHistEscolar.Update
        Next
    End With
End Sub




Private Sub Cb_Deficiencia_DropDown()
    Cb_Deficiencia.Clear
    Set RsDeficiencia = BD.OpenRecordset("SELECT * FROM Deficiencias ORDER BY Descr")
    If RsDeficiencia.BOF And RsDeficiencia.EOF Then
        Else
            RsDeficiencia.MoveFirst
            Do Until RsDeficiencia.EOF
                Cb_Deficiencia.AddItem (RsDeficiencia.Fields("Descr"))
                RsDeficiencia.MoveNext
            Loop
    End If
End Sub


Private Sub Cb_Nome_Change()
    If Len(Trim(Cb_Nome.Text)) >= 50 Then
        Cb_Nome.Text = Mid(Trim(Cb_Nome.Text), 1, 50)
        Beep
    End If
   'Cb_Nome_DropDown
End Sub

Private Sub Cb_Nome_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
       CarregarMatricula (formBuscar.IniciarBusca("Matriculas"))
    End If

End Sub
Private Sub CarregarMatricula(Matr As String)
        Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID ='" & Matr & "'")
        With RsMatricula
            If .BOF And .EOF Then
                    MsgBox "Matricula não encontrada.", vbInformation, "CESNet - Aviso!"
                    LimpDados
                    SST_Geral.Tab = 0
                    MebMatricula.SetFocus
                    Exit Sub
                Else
                    MstDadosAluno
                    LstHstEscolar
            End If
        End With
    'End If


End Sub
Private Sub Cb_Nome_LostFocus()
    Dim RsTMP As Recordset
    If Acao <> 1 Then Exit Sub
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Matriculas WHERE Nome = '" & Cb_Nome.Text & "'")
    If RsTMP.BOF And RsTMP.EOF Then
        RsTMP.Close
            Exit Sub
        Else
            RsTMP.MoveFirst
            If MsgBox("Já existe uma matricula com esse nome. Deseja carregar os dados?", vbInformation + vbYesNo, "CESNet - Aviso") = vbYes Then
                MatrID = RsTMP.Fields("MatrID")
                MstDadosAluno
            End If
    End If
    
End Sub


Private Sub Cb_SerieHistEsc_DropDown()
    Cb_SerieHistEsc.Clear
    Set RsSerie = BD.OpenRecordset("SELECT * FROM Serie ORDER BY Descr")
    If RsSerie.BOF And RsSerie.EOF Then
            MsgBox "Não existem series cadastradas. Por favor cadastre." & Chr(13) & Chr(13) & "Operação cancelada...", vbInformation, "CESNet - Aviso!"
            Bt_CancelarHistEsc_Click
        Else
        RsSerie.MoveFirst
        Do Until RsSerie.EOF
            Cb_SerieHistEsc.AddItem (RsSerie.Fields("Descr"))
            RsSerie.MoveNext
        Loop
    End If
End Sub

Private Sub Cb_SerieHistEsc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub


Private Sub Cb_Unidade_Click()
    Set RsUnidade = BD.OpenRecordset("SELECT * FROM Unidades WHERE UnidID = '" & left(CB_Unidade.Text, 3) & "'")
    If RsUnidade.BOF And RsUnidade.EOF Then
            MsgBox "Erro ao localizar Unidade de Ensino!", vbInformation, "CESNet - Aviso"
            RsUnidade.Close
            Exit Sub
        Else
            RsUnidade.MoveFirst
            Txt_Mun.Text = IIf(IsNull(RsUnidade.Fields("Mun")), " ", UCase(RsUnidade.Fields("Mun")))
            CB_UF.Text = IIf(IsNull(RsUnidade.Fields("UF")), " ", UCase(RsUnidade.Fields("UF")))
            RsUnidade.Close
    End If
End Sub

Private Sub CB_Unidade_GotFocus()
    CB_Unidade.Clear
    Set RsUnidade = BD.OpenRecordset("SELECT * FROM Unidades ORDER BY UnidID")
    With RsUnidade
        If .BOF And .EOF Then
                tmp = MsgBox("Antes de cadastrar ou alterar qualquer aluno é necessário cadastrar uma Unidade de Ensino!" & Chr(13) & "Gostaria de Cadastrar agora?", vbYesNo, "CESNet - Aviso")
                If tmp = 6 Then
                        Unload Me
                        Form_Unidade.Show
                        Exit Sub
                    Else
                        Exit Sub
                End If
            Else
                .MoveFirst
                Do While .EOF = False
                    CB_Unidade.AddItem (.Fields("UnidID") & " - " & .Fields("Nome"))
                    .MoveNext
                Loop
                .Close
                CB_Unidade.Text = UnidadeEnsino & " - " & UnidadeEnsinoNome
        End If
    End With

End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub




Private Sub MebMatricula_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim matrtmp As String
    If KeyCode = 114 Then
        matrtmp = Trim(formBuscar.IniciarBusca("Matriculas"))
        If matrtmp = "" Or matrtmp = 0 Then Exit Sub
        CarregarMatricula (matrtmp)
    End If
End Sub

Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            Acao = 1
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            HDFormDados (True)
            hdMenu (False)
            'LimpaMatricula
            '-------No BOTAO-------------------------
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub

            LimpDados
            HDFormDados (True)
            Cb_Nome.Clear
            Meb_Dt.Text = Format(Date, "dd/mm/yyyy")

            CB_UF.Text = "RJ"
            Txt_Nac.Text = "BRASILEIRA"
            MatrID = ""
            Cb_Nome.SetFocus
    
            'Acao = 1

        Case 2 'Alterar
            Acao = 2
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            HDFormDados (True)
            hdMenu (False)
            '-------- NO BOTAO ------------------
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    
            MebMatricula.PromptInclude = False
            If MebMatricula.Text = "" Then
                MebMatricula.PromptInclude = True
                Exit Sub
            End If
            MebMatricula.PromptInclude = True
            HDFormDados (True)
    
    
            CB_Unidade.Enabled = False
    

            
        Case 3 'Excluir
            Acao = 3
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            ExcluirDados
            
            
        Case 4 'Imprimir
            Acao = 4
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            'ImprimirListagem
        Case 6 'Matricula
            Form_Matricula.CarregarMatricula (MatrID)
        Case 8 'Gravar
            'GravarDados
            'LimpaMatricula
            'matriculas
            If ValidarSoftware("Matriculas") = False Then Exit Sub
            'MstDados
            '------------ NO BOTAO -------------
            If GrvDados = False Then Exit Sub
            HDFormDados (False)
            hdMenu (True)
        Case 9 'Cancelar
            Acao = 7
            HDFormDados (False)
            hdMenu (True)
            'LimpaMatricula
            '----NO BOTAO ----
            Acao = 0
            HDFormDados (False)
            LimpDados
            CB_Unidade.Clear
            LstNomeAlunos
            'Chk_Retorno.Value = 0
            'Chk_Retorno.Enabled = False
            Meb_Dt.Enabled = False
            With RsMatricula
                If .BOF And .EOF Then
                    Else
                        RsMatricula.MoveFirst
                        MstDadosAluno
                End If
            End With
        Case 11 'eXIBIR FOTO
            If ChkAcesso(Me.Name, "C") = False Then Exit Sub
            Form_ExibirImagem.ExibirFoto (MatrID)
        Case 13 'CAPTURAR FOTO
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            If MatrID = "" Then
                    MsgBox "Favor selecionar uma Matrícula!", vbInformation, "CESNet - Aviso"
                    Exit Sub
                Else
                    frmMain.InicializarCaptura (MatrID)
            End If
    End Select
End Sub
Private Sub hdMenu(op As Boolean)

    Tb_Menu.Buttons(1).Enabled = op
    Tb_Menu.Buttons(2).Enabled = op
    Tb_Menu.Buttons(3).Enabled = op
    Tb_Menu.Buttons(4).Enabled = op
    Tb_Menu.Buttons(6).Enabled = op
    
    Tb_Menu.Buttons(8).Enabled = IIf(op = False, True, False)
    Tb_Menu.Buttons(9).Enabled = IIf(op = False, True, False)
    
End Sub




Private Sub Bt_AltObs_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub

    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    
    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Then
        MebMatricula.PromptInclude = True
        Exit Sub
    End If
    MebMatricula.PromptInclude = True
    Txt_Obs.Enabled = True
    'DTP_ValidadeCard.Enabled = True
    Bt_GrvObs.Enabled = True
    Bt_AltObs.Enabled = False
End Sub



Private Sub ExcluirDados()
    On Error GoTo TratErro
    If ChkAcesso(Me.Name, "E") = False Then Exit Sub

    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Then
        MebMatricula.PromptInclude = True
        Exit Sub
    End If
    'Autentica o usuario
    If Form_AutenticacaoUsuario.CarregarForm = False Then
        Exit Sub
    End If

    MebMatricula.PromptInclude = True
    With RsMatricula
        .FindFirst "MatrID = '" & MatrID & "'"
        If .NoMatch Then
                MsgBox "Registro Não Encontrado. Tente Novamente!", vbExclamation, "CESNet - Aviso"
                Exit Sub
            Else
                'Checa se existem modulos emprestados
                Set RsTMP = BD.OpenRecordset("SELECT * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "' AND IsNull(DtDevolucao)")
                If RsTMP.BOF And RsTMP.EOF Then
                    Else
                        MsgBox "Esta matricula possui MÓDULO(S) emprestado(s) caso a exclua essa matricula tal informação irá desaparecer!", vbInformation, "CESNet - Atenção"
                End If
                'Checa se existem livros emprestados
                Set RsTMP = BD.OpenRecordset("SELECT * FROM BibliotecaEmprestimo WHERE MatrID = '" & MatrID & "' AND IsNull(DtDevolucao)")
                If RsTMP.BOF And RsTMP.EOF Then
                    Else
                        MsgBox "Esta matricula possui LIVRO(S) emprestado(s) caso a exclua essa matricula tal informação irá desaparecer!", vbInformation, "CESNet - Atenção"
                End If
                tmp = MsgBox("Deseja EXCLUIR esta matricula " & MatrID & " e todo seu HISTÓRICO?", vbInformation + vbYesNo, "CESNet - EXCLUIR")
                If tmp = 6 Then
                BD.Execute "DELETE * FROM Matriculas WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM ProvasTMP WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM HistEscolar WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM BibliotecaEmprestimo WHERE MatrID = '" & MatrID & "'"
                BD.Execute "DELETE * FROM EmprestimoModulo WHERE MatrID = '" & MatrID & "'"
                Kill PathBD & "\Database\Img\" & Format(MatrID, "000000000") & "001.jpg"
                Call RegLog(MatrID, "excluiu a matricula.")
                Acao = 0
                HDFormDados (False)
                LimpDados
                CB_Unidade.Clear
                Meb_Dt.Enabled = False
            End If
        End If
    End With
    Exit Sub
TratErro:
    Call RegLogErros(Err.Number, Err.Description, Me.Caption, UsuarioID)
    If Err.Number = 53 Then Resume Next
    MsgBox "Erro ao excluir Matricula" & Chr(13) & "Descrição: " & Err.Description, vbInformation, Err.Number
End Sub



Private Sub Bt_GrvObs_Click()
On Error GoTo TrtErro
    With RsMatricula
        .FindFirst "MatrID = '" & MebMatricula.Text & "'"
        If .NoMatch Then
                MsgBox "Erro no acesso ao Banco de Dados." & Chr(13) & "Por favor, reinicie o formulário!", vbExclamation, "aviso!"
            Else
                .Edit
                .Fields("Obs") = Txt_Obs.Text
                
                .Update
                Txt_Obs.Enabled = False
                Bt_GrvObs.Enabled = False
                Bt_AltObs.Enabled = True
        End If
    End With
    Exit Sub
TrtErro:
    MsgBox Err.Description, vbInformation, Err.Number
    Resume Next
End Sub
Private Function GrvDados() As Boolean
'On Error GoTo TratErroGrv
    If Cb_Nome.Text = "" Or CB_Unidade.Text = "" Then
        MsgBox "Os campos: NOME OU UNIDADE, não podem ser deixados em branco." & Chr(13) & "Por favor Verifique!", vbExclamation, "Aviso!"
        Cb_Nome.SetFocus
        GrvDados = False
        Exit Function
    End If
    Meb_Nasc.PromptInclude = False
    If Trim(Meb_Nasc.Text) = "" Then
            Meb_Nasc.PromptInclude = True
        Else
            Meb_Nasc.PromptInclude = True
            If IsDate(Meb_Nasc.Text) Then
                Else
                    MsgBox "Data de nascimento não é valida, por favor verifique", vbInformation, "CESNet - Atenção"
                    GrvDados = False
                    Exit Function
            End If
    End If
    'If Trim(Txt_End.Text) = "" Or Trim(Txt_Num.Text) = "" Then
   '     MsgBox "O campo NÚMERO do endereço nao pode ser um valor vazio!", vbInformation, "CESNet - Aviso"
   '     Txt_Num.SetFocus
   '     Exit Sub
   ' End If
    
    
    DoEvents
    HDFormDados (False)
    Meb_Dt.Enabled = False
    
    'Dim MatrID As String
    MatrID = Right(Meb_Dt.Text, 2) & "." & left(CB_Unidade.Text, 3)
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID LIKE '" & MatrID & "*'")
    With RsTMP
        If .BOF And .EOF Then
            'If Chk_Retorno.Value = 0 Then
                        MatrID = MatrID & ".0001"
                    'Else
                     '   MatrID = MatrID & "." & Meb_MatrAnt.Text
                'End If
            Else
                .MoveLast
                'If Chk_Retorno.Value = 0 Then
                        MatrID = MatrID & "." & Mid(String(4, "0"), 1, 4 - Len(Right(.Fields("MatrID") + 1, 4))) & Right(.Fields("MatrID") + 1, 4)
                    'Else
                        'If Trim(Meb_MatrAnt.Text) = "" Or Trim(Meb_MatrAnt.Text) = "0000" Then
                        '    MsgBox "Favor informar o número da matricula de retorno!", vbInformation, "CESNet - Aviso"
                        '    Exit Sub
                        'End If
                        '.FindFirst "MatrID = '" & MatrID & "." & Meb_MatrAnt.Text & "'"
                        'If .NoMatch Then
                        '        MatrID = MatrID & "." & Mid(String(4, "0"), 1, 4 - Len(Right(Meb_MatrAnt.Text, 4))) & Right(Meb_MatrAnt.Text, 4)
                        '    Else
                        '        MsgBox "Matricula já cadastrada. Por favor Verifique!", vbExclamation, "Aviso!"
                        '        Meb_MatrAnt.SetFocus
                        '        Exit Sub
                        'End If
                'End If
        End If
    End With
    With RsMatricula
        
        Select Case Acao
            Case 1 'INCLUIR MATRICULA
                .FindFirst "Nome = '" & rc(Trim(Cb_Nome.Text)) & "'"
                If .NoMatch = False Then
                    MsgBox "Nome de aluno já cadastrado no sistema. Operação cancelada." & Chr(13) & "Por favor verifique!", vbInformation, "CESNet - Atenção!"
                    Exit Function
                End If
                'tmp = .Fields("MatrID")
                .AddNew
                .Fields("MatrID") = MatrID
                .Fields("Nome") = Trim(Cb_Nome.Text)
                .Fields("UnidadeID") = left(CB_Unidade.Text, 3)
                .Fields("Unidade") = Trim(Mid(CB_Unidade.Text, 6, Len(CB_Unidade.Text)))
                .Fields("DtRetorno") = Null
                .Fields("DtMat") = Meb_Dt.Text
                .Fields("End") = IIf(Trim(Txt_End.Text) = "", Null, Trim(Txt_End.Text))
                .Fields("Numero") = IIf(Trim(Txt_Num.Text) = "", Null, Trim(Txt_Num.Text))
                .Fields("Compl") = IIf(Trim(Txt_Comp.Text) = "", Null, Trim(Txt_Comp.Text))
                .Fields("Bai") = IIf(Trim(Txt_Bai.Text) = "", Null, Trim(Txt_Bai.Text))
                .Fields("Mun") = IIf(Trim(Txt_Mun.Text) = "", Null, Trim(Txt_Mun.Text))
                .Fields("UF") = IIf(Trim(CB_UF.Text) = "", Null, Trim(CB_UF.Text))
                .Fields("CEP") = IIf(Trim(Meb_Cep.Text) = "", Null, Trim(Meb_Cep.Text))
                .Fields("Sexo") = IIf(Trim(Cb_Sexo.Text) = "", Null, Trim(Cb_Sexo.Text))
                Meb_Nasc.PromptInclude = False
                If Trim(Meb_Nasc.Text) = "" Then
                    Else
                        Meb_Nasc.PromptInclude = True
                        .Fields("Nasc") = Trim(Meb_Nasc.Text)
                End If
                    
                Meb_Nasc.PromptInclude = True
                .Fields("Mail") = IIf(Trim(Txt_Mail.Text) = "", Null, Trim(Txt_Mail.Text))
                .Fields("Cel") = IIf(Trim(Txt_Cel.Text) = "", Null, Trim(Txt_Cel.Text))
                .Fields("Tel1") = IIf(Trim(Txt_Tel1.Text) = "", Null, Trim(Txt_Tel1.Text))
                .Fields("Tel2") = IIf(Trim(Txt_Tel2.Text) = "", Null, Trim(Txt_Tel2.Text))
    
                .Fields("RG") = IIf(Trim(Txt_RG.Text) = "", Null, Trim(Txt_RG.Text))
                .Fields("OE") = IIf(Trim(Txt_OrgEmi.Text) = "", Null, Trim(Txt_OrgEmi.Text))
                .Fields("CPF") = IIf(Trim(Txt_CPF.Text) = "", Null, Trim(Txt_CPF.Text))
                .Fields("CertNasc") = IIf(Trim(Txt_CertNasc.Text) = "", Null, Trim(Txt_CertNasc.Text))
                .Fields("Natural") = IIf(Trim(Txt_Nat.Text) = "", Null, Trim(Txt_Nat.Text))
                .Fields("NaturalUF") = IIf(Trim(CB_UFNatural.Text) = "", Null, Trim(CB_UFNatural.Text))
                .Fields("EstCivil") = IIf(Trim(CB_EstCiv.Text) = "", Null, Trim(CB_EstCiv.Text))
                .Fields("Nacion") = IIf(Trim(Txt_Nac.Text) = "", Null, Trim(Txt_Nac.Text))
                    
                .Fields("DefID") = IIf(Trim(Cb_Deficiencia.Text) = "", Null, PgIDDef(Trim(Cb_Deficiencia.Text)))
                    
                .Fields("Mae") = IIf(Trim(Txt_Mae.Text) = "", Null, Trim(Txt_Mae.Text))
                mebDtNascMae.PromptInclude = False
                .Fields("DtNascMae") = IIf(Trim(mebDtNascMae.Text) = "", Null, Format(Trim(mebDtNascMae.Text), "##/##/####"))
                mebDtNascMae.PromptInclude = True
                
                mebDtNascPai.PromptInclude = False
                .Fields("Pai") = IIf(Trim(Txt_Pai.Text) = "", Null, Trim(Txt_Pai.Text))
                .Fields("DtNascPai") = IIf(Trim(mebDtNascPai.Text) = "", Null, Format(Trim(mebDtNascPai.Text), "##/##/####"))
                mebDtNascPai.PromptInclude = True
                
                .Fields("Raca") = IIf(Trim(txtRaca.Text) = "", Null, Trim(txtRaca.Text))
                .Fields("OpcaoRel") = IIf(Trim(txtOpcaoRel.Text) = "", Null, Trim(txtOpcaoRel.Text))
                .Fields("TpSang") = IIf(Trim(txtTpSang.Text) = "", Null, Trim(txtTpSang.Text))
                
                .Fields("UsuarioID") = UsuarioID
                .Fields("DtHrSis") = Now
                
                .Update
                MebMatricula.Text = MatrID
                MsgBox "Nova Matricula: " & MatrID, vbDefaultButton1, "CESNet - Nova Matricula"
                Call RegLog(MatrID, "Inclusao de DADOS PESSOAIS do aluno.")
                
            Case 2 'ALTERAR MATRICULA
                .FindFirst "MatrID = '" & MebMatricula.Text & "'"
                If .NoMatch = True Then
                    MsgBox "Erro ao localizar matricula: " & MatrID & ". Caso o problema continue chame o suporte.", vbInformation, "CESNet - Atenção"
                    GrvDados = False
                    Exit Function
                End If
                .Edit
                .Fields("Nome") = Trim(Cb_Nome.Text)
                .Fields("UnidadeID") = left(Trim(CB_Unidade.Text), 3)
                .Fields("Unidade") = Trim(Mid(CB_Unidade.Text, 6, Len(CB_Unidade.Text)))
                
                
                
                .Fields("End") = IIf(Trim(Txt_End.Text) = "", Null, Trim(Txt_End.Text))
                .Fields("Numero") = IIf(Trim(Txt_Num.Text) = "", Null, Trim(Txt_Num.Text))
                .Fields("Compl") = IIf(Trim(Txt_Comp.Text) = "", Null, Trim(Txt_Comp.Text))
                .Fields("Bai") = IIf(Trim(Txt_Bai.Text) = "", Null, Trim(Txt_Bai.Text))
                .Fields("Mun") = IIf(Trim(Txt_Mun.Text) = "", Null, Trim(Txt_Mun.Text))
                .Fields("UF") = IIf(Trim(CB_UF.Text) = "", Null, Trim(CB_UF.Text))
                .Fields("CEP") = IIf(Trim(Meb_Cep.Text) = "", Null, Trim(Meb_Cep.Text))
                .Fields("Sexo") = IIf(Trim(Cb_Sexo.Text) = "", Null, Trim(Cb_Sexo.Text))
                Meb_Nasc.PromptInclude = False
                If Trim(Meb_Nasc.Text) = "" Then
                    Else
                        Meb_Nasc.PromptInclude = True
                        .Fields("Nasc") = Meb_Nasc.Text
                End If
                Meb_Nasc.PromptInclude = True
                .Fields("Mail") = IIf(Trim(Txt_Mail.Text) = "", Null, Trim(Txt_Mail.Text))
                .Fields("Cel") = IIf(Trim(Txt_Cel.Text) = "", Null, Trim(Txt_Cel.Text))
                .Fields("Tel1") = IIf(Trim(Txt_Tel1.Text) = "", Null, Trim(Txt_Tel1.Text))
                .Fields("Tel2") = IIf(Trim(Txt_Tel2.Text) = "", Null, Trim(Txt_Tel2.Text))
    
                .Fields("RG") = IIf(Trim(Txt_RG.Text) = "", Null, Trim(Txt_RG.Text))
                .Fields("OE") = IIf(Trim(Txt_OrgEmi.Text) = "", Null, Trim(Txt_OrgEmi.Text))
                .Fields("CPF") = IIf(Trim(Txt_CPF.Text) = "", Null, Trim(Txt_CPF.Text))
                .Fields("CertNasc") = IIf(Trim(Txt_CertNasc.Text) = "", Null, Trim(Txt_CertNasc.Text))
                .Fields("Natural") = IIf(Trim(Txt_Nat.Text) = "", Null, Trim(Txt_Nat.Text))
                .Fields("NaturalUF") = IIf(Trim(CB_UFNatural.Text) = "", Null, Trim(CB_UFNatural.Text))
                .Fields("EstCivil") = IIf(Trim(CB_EstCiv.Text) = "", Null, Trim(CB_EstCiv.Text))
                .Fields("Nacion") = IIf(Trim(Txt_Nac.Text) = "", Null, Trim(Txt_Nac.Text))
                    
                .Fields("DefID") = IIf(Trim(Cb_Deficiencia.Text) = "", Null, PgIDDef(Trim(Cb_Deficiencia.Text)))
                    
                .Fields("Mae") = IIf(Trim(Txt_Mae.Text) = "", Null, Trim(Txt_Mae.Text))
                .Fields("Pai") = IIf(Trim(Txt_Pai.Text) = "", Null, Trim(Txt_Pai.Text))
                
                mebDtNascPai.PromptInclude = False
                mebDtNascMae.PromptInclude = False
                
                .Fields("DtNascMae") = IIf(Trim(mebDtNascMae.Text) = "", Null, Format(Trim(mebDtNascMae.Text), "##/##/####"))
                .Fields("DtNascPai") = IIf(Trim(mebDtNascPai.Text) = "", Null, Format(Trim(mebDtNascPai.Text), "##/##/####"))
                
                mebDtNascPai.PromptInclude = True
                mebDtNascMae.PromptInclude = True
                
                .Fields("Raca") = IIf(Trim(txtRaca.Text) = "", Null, Trim(txtRaca.Text))
                .Fields("OpcaoRel") = IIf(Trim(txtOpcaoRel.Text) = "", Null, Trim(txtOpcaoRel.Text))
                .Fields("TpSang") = IIf(Trim(txtTpSang.Text) = "", Null, Trim(txtTpSang.Text))
                .Fields("UsuarioID") = UsuarioID
                .Fields("DtHrSis") = Now
                .Update
                MatrID = MebMatricula.Text
                MsgBox "Matricula: " & MatrID & " alterada com sucesso!", vbDefaultButton1, "CESNet - Nova Matricula"
                Call RegLog(MatrID, "Alterou os dados pessoais do aluno.")
        End Select
        Acao = 0
        GrvDados = True

        'MatrID = MatrID
    End With
    Exit Function
TratErroGrv:
    RegLogErros Err.Number, Err.Description, Me.Caption, UsuarioID
    MsgBox "Descrição: " & Err.Description & Chr(13) & "Dados não gravados...", vbInformation, "Erro n. " & Err.Number
    GrvDados = False
End Function






Private Sub Form_Load()
    SST_Geral.Tab = 0
    Me.top = 0
    Me.left = 0
    Acao = 0
    Meb_Dt.Text = Format(Date, "dd/MM/yyyy")
    MatrID = ""
    'TodasDisciplinas = 0
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas ORDER BY Nome")
    hdMenu (True)

End Sub
Private Sub Cb_Nome_Click()

    If Cb_Nome.Text = "" Then Exit Sub
    With RsMatricula
        .FindFirst "Nome ='" & rc(Cb_Nome.Text) & "'"
        If .NoMatch Then
                MsgBox "Erro ao localizar nome do aluno no banco de dados." & Chr(13) & "Por favor, reinicie o formulário!", vbExclamation, "aviso!"
                Exit Sub
            Else
                MstDadosAluno
                LstHstEscolar
        End If
    End With
End Sub
Private Sub Cb_Nome_DropDown()
    If Acao = 1 Or Acao = 2 Then Exit Sub
    Set RsMatricula = BD.OpenRecordset("SELECT TOP 100 * FROM Matriculas WHERE Nome LIKE '" & rc(Cb_Nome.Text) & "*' ORDER BY Nome")
    If RsMatricula.BOF And RsMatricula.EOF Then
            Cb_Nome.Clear
        Else
            LstNomeAlunos
    End If
End Sub

Private Sub Cb_Nome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
    If Len(Cb_Nome.Text) >= 50 Then
        KeyAscii = 0
        Beep
    End If
End Sub
Private Sub LstNomeAlunos()
    With RsMatricula
        If .BOF And .EOF Then
                Exit Sub
            Else
                .MoveFirst
                'If Trim(Cb_Nome.Text) = "" Then Exit Sub
                tmp = Cb_Nome.Text
                Cb_Nome.Clear
                Cb_Nome.Text = tmp
                Do While .EOF = False
                    Cb_Nome.AddItem (IIf(IsNull(.Fields("Nome")), "", .Fields("Nome")))
                    .MoveNext
                Loop
        End If
    End With
End Sub

Private Sub MstDadosAluno()
    Dim UnidTMP As String
    With RsMatricula
        If Trim(.Fields("MatrID")) = "" Or IsNull(.Fields("MatrID")) Then
            Exit Sub
        End If
        Meb_Nasc.PromptInclude = False
        MatrID = .Fields("MatrID")
        MebMatricula.Text = MatrID
        

        Meb_Dt.Text = PgDadosMatr(MatrID).DtMatr
        Cb_Nome.Text = PgDadosMatr(MatrID).Nome
        If PgDadosMatr(MatrID).UnidMatrID = "" Then
                UnidTMP = PgDadosMatr(MatrID).UnidMatr
            Else
                UnidTMP = PgDadosMatr(MatrID).UnidMatrID & " - " & PgDadosMatr(MatrID).UnidMatr
        End If
        CB_Unidade.Clear
        CB_Unidade.AddItem IIf(UnidTMP = "", " ", UnidTMP)
        CB_Unidade.Text = CB_Unidade.List(0)
        
        Txt_End.Text = PgDadosMatr(MatrID).Endereco
        Txt_Num.Text = PgDadosMatr(MatrID).Numero
        Txt_Comp.Text = PgDadosMatr(MatrID).Compl
        Txt_Bai.Text = PgDadosMatr(MatrID).Bairro
        Txt_Mun.Text = PgDadosMatr(MatrID).Munic
        CB_UF.Text = PgDadosMatr(MatrID).UF
        Meb_Cep.Text = PgDadosMatr(MatrID).CEP
        Cb_Sexo.Text = PgDadosMatr(MatrID).Sexo
        
        Meb_Nasc.Text = PgDadosMatr(MatrID).Nasc
        
        Txt_Mail.Text = PgDadosMatr(MatrID).Mail
        Txt_Cel.Text = PgDadosMatr(MatrID).Cel
        Txt_Tel1.Text = PgDadosMatr(MatrID).Tel1
        Txt_Tel2.Text = PgDadosMatr(MatrID).Tel2
        Txt_RG.Text = PgDadosMatr(MatrID).RG
        Txt_OrgEmi.Text = PgDadosMatr(MatrID).OE
        Txt_CPF.Text = PgDadosMatr(MatrID).CPF
        Txt_CertNasc.Text = PgDadosMatr(MatrID).CertNasc
        Txt_Nat.Text = PgDadosMatr(MatrID).Natural
        CB_UFNatural.Text = PgDadosMatr(MatrID).NaturalUF
        CB_EstCiv.Text = PgDadosMatr(MatrID).EstCivil
        Txt_Nac.Text = PgDadosMatr(MatrID).Nacion
        
        Txt_Mae.Text = PgDadosMatr(MatrID).Mae
        Txt_Pai.Text = PgDadosMatr(MatrID).Pai
        mebDtNascMae.PromptInclude = False
        mebDtNascMae.Text = PgDadosMatr(MatrID).DtNascMae
        mebDtNascMae.PromptInclude = True
        mebDtNascPai.PromptInclude = False
        mebDtNascPai.Text = PgDadosMatr(MatrID).DtNascPai
        mebDtNascPai.PromptInclude = True
        
        
        txtRaca.Text = PgDadosMatr(MatrID).Raca
        txtOpcaoRel.Text = PgDadosMatr(MatrID).OpcaoRel
        txtTpSang.Text = PgDadosMatr(MatrID).TpSang
        
        
        Cb_Deficiencia.AddItem (PgNomeDef(PgDadosMatr(MatrID).Deficiencia))
        Cb_Deficiencia.Text = PgNomeDef(PgDadosMatr(MatrID).Deficiencia)
        
        Txt_Obs.Text = PgDadosMatr(MatrID).Obs
        Meb_Nasc.PromptInclude = True
    End With
    SST_Geral.Tab = 0
    SST_DadosPessoais.Tab = 0
    '***** Checar Aviso ******
    If PgAviso(MatrID) = True Then
        Exit Sub
    End If
    '*************************

End Sub
Private Sub LimpDados()
    MebMatricula.PromptInclude = False
    Meb_Nasc.PromptInclude = False
    
    MebMatricula.Text = ""
    Cb_Nome.Text = ""
    CB_Unidade.Clear
    Txt_End.Text = ""
    Txt_Num.Text = ""
    Txt_Comp.Text = ""
    Txt_Bai.Text = ""
    Txt_Mun.Text = ""
    CB_UF.Text = CB_UF.List(0)
    Meb_Cep.Text = ""
    Cb_Sexo.Text = Cb_Sexo.List(0)
    Meb_Nasc.Text = ""
    
    Txt_Mail.Text = ""
    Txt_Cel.Text = ""
    Txt_Tel1.Text = ""
    Txt_Tel2.Text = ""
    
    Txt_RG.Text = ""
    Txt_OrgEmi.Text = ""
    Txt_CPF.Text = ""
    Txt_CertNasc.Text = ""
    Txt_Nat.Text = ""
    CB_UFNatural.Text = CB_UFNatural.List(0)
    CB_EstCiv.Text = CB_EstCiv.List(0)
    Txt_Nac.Text = ""
    
    Cb_Deficiencia.Clear
    
    Txt_Mae.Text = ""
    Txt_Pai.Text = ""
    mebDtNascMae.PromptInclude = False
    mebDtNascMae.Text = ""
    mebDtNascMae.PromptInclude = True
    
    mebDtNascPai.PromptInclude = False
    mebDtNascPai.Text = ""
    mebDtNascPai.PromptInclude = True
    
    txtRaca.Text = ""
    txtOpcaoRel.Text = ""
    txtTpSang.Text = ""
    
    MebMatricula.PromptInclude = True
    Meb_Nasc.PromptInclude = True

End Sub


Private Sub HDFormDados(op As Boolean)
    MebMatricula.Enabled = IIf(op = True, False, True)
    CB_Unidade.Enabled = op
    Txt_End.Enabled = op
    Txt_Num.Enabled = op
    Txt_Comp.Enabled = op
    Txt_Bai.Enabled = op
    Txt_Mun.Enabled = op
    CB_UF.Enabled = op
    Cb_Sexo.Enabled = op
    Meb_Nasc.Enabled = op
    Meb_Cep.Enabled = op
    
    Txt_Mail.Enabled = op
    Txt_Cel.Enabled = op
    Txt_Tel1.Enabled = op
    Txt_Tel2.Enabled = op
    
    Txt_RG.Enabled = op
    Txt_OrgEmi.Enabled = op
    Txt_CPF.Enabled = op
    Txt_CertNasc.Enabled = op
    Txt_Nat.Enabled = op
    CB_UFNatural.Enabled = op
    CB_EstCiv.Enabled = op
    Txt_Nac.Enabled = op
    
    Cb_Deficiencia.Enabled = op
    
    Txt_Mae.Enabled = op
    Txt_Pai.Enabled = op
    mebDtNascMae.Enabled = op
    mebDtNascPai.Enabled = op
    
    txtRaca.Enabled = op
    txtOpcaoRel.Enabled = op
    txtTpSang.Enabled = op
End Sub



Private Sub Meb_Cep_LostFocus()
    If Acao = 1 Then PgEndereco (Meb_Cep.Text)
End Sub

Private Sub Meb_Dt_GotFocus()
    Meb_Dt.SelStart = 0
    Meb_Dt.SelLength = 10
End Sub
Private Sub Meb_Dt_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Cb_Unidade_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub
Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID ='" & MebMatricula.Text & "'")
        With RsMatricula
            If .BOF And .EOF Then
                    MsgBox "Matricula não encontrada.", vbInformation, "CESNet - Aviso!"
                    LimpDados
                    SST_Geral.Tab = 0
                    MebMatricula.SetFocus
                    Exit Sub
                Else
                    MstDadosAluno
                    LstHstEscolar
            End If
        End With
    End If
End Sub


Private Sub MSFG_HistEscolar_DblClick()
    If MatrID = "" Then Exit Sub
    With MSFG_HistEscolar
        Cb_SerieHistEsc.AddItem (IIf(Trim(.TextMatrix(.Row, 0)) = "", " ", .TextMatrix(.Row, 0)))
        Cb_SerieHistEsc.Text = IIf(Trim(.TextMatrix(.Row, 0)) = "", " ", .TextMatrix(.Row, 0))
        Txt_AnoHistEsc.Text = .TextMatrix(.Row, 1)
        Txt_EstabHistEsc.Text = .TextMatrix(.Row, 2)
        Txt_CidadeHistEsc.Text = .TextMatrix(.Row, 3)
                
        .Enabled = False
        
        Cb_SerieHistEsc.Enabled = True
        Txt_AnoHistEsc.Enabled = True
        Txt_EstabHistEsc.Enabled = True
        Txt_CidadeHistEsc.Enabled = True
        
        'Bt_ImprCert.Enabled = False
        Bt_GravarHistEsc.Enabled = True
        Bt_CancelarHistEsc.Enabled = True
        
    End With
End Sub
Private Sub MSFG_HistEscolar_KeyDown(KeyCode As Integer, Shift As Integer)
    Select Case KeyCode
        Case 45
              MSFG_HistEscolar.Rows = MSFG_HistEscolar.Rows + 1
            
        Case 46
            If MSFG_HistEscolar.Rows = 2 Then
                    MSFG_HistEscolar.Rows = 1
                    MSFG_HistEscolar.Rows = 2
                Else
                    MSFG_HistEscolar.RemoveItem (MSFG_HistEscolar.Row)
            End If
    End Select
End Sub
Private Sub SST_Geral_Click(PreviousTab As Integer)
    If Acao <> 0 Then SST_Geral.Tab = 0
    If SST_Geral.Tab = 0 Then
        SST_DadosPessoais.Tab = 0
    End If
    If SST_Geral.Tab = 1 Then
    End If
End Sub
Private Sub Tb_Menu_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim strSQL As String
    If MatrID = "" Then Exit Sub
    Select Case ButtonMenu.Parent.Key & ButtonMenu.Index
        Case 1 'Ficha de Matricula
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            
            strSQL = "SELECT * FROM Matriculas WHERE MatrID = '" & MatrID & "'"
            Call Relatorio(rptFichaCadastro, strSQL)
            
            If ChkExisteArquivo(PathBD & "\Database\IMG\" & Format(MatrID, "000000000") & "001.jpg") = True Then
                Set rptFichaCadastro.Sections("Corpo").Controls.Item("imgFoto").Picture = LoadPicture(PathBD & "\Database\IMG\" & Format(MatrID, "000000000") & "001.jpg")
            End If
            rptFichaCadastro.Sections("Corpo").Controls.Item("lblCurso").Caption = PgNomeEnsino(PgMatrEnsino(MatrID))
            
            'Set rptFichaCadastro.Sections("Corpo").Controls.Item("lbLocal").Caption = PgDadosUnid(UnidadeEnsino).Municipio
            'Carteira
            If CartConjugada = 1 Then
                    rptFichaCadastro.Sections("Rodape").Visible = True
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblUnidade").Caption = PgDadosUnid(UnidadeEnsino).Nome
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lbUA").Caption = "U.A.: " & PgDadosUnid(UnidadeEnsino).UA
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblMatr").Caption = MatrID
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblnome").Caption = PgDadosMatr(MatrID).Nome
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblNasc").Caption = PgDadosMatr(MatrID).Nasc
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblVenc").Caption = PgDadosMatr(MatrID).ValCard
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblPai").Caption = PgDadosMatr(MatrID).Pai
                    rptFichaCadastro.Sections("Rodape").Controls.Item("lblMae").Caption = PgDadosMatr(MatrID).Mae
                Else
                    rptFichaCadastro.Sections("Rodape").Visible = False
                
            End If
            rptFichaCadastro.Show 1
            
            Exit Sub
            'MebMatricula.PromptInclude = False
            'If MebMatricula.Text = "" Then
            '        MebMatricula.PromptInclude = True
            '        Exit Sub
            '    Else
            '        MebMatricula.PromptInclude = True
            '        '*************************************************
            '        If Form_Impressora.LoadFormCI(True, True, False, False, False, True, True, True, False, False) = False Then
            '            Exit Sub
            '        End If
            '        Call ImpFichaMatr(MebMatricula.Text)
            'End If
        Case 2 'Carteira
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            strSQL = "SELECT * FROM Matriculas WHERE MatrID = '" & MatrID & "'"
            rptCarteiraID.Sections("Corpo").Controls.Item("lblUnidade").Caption = PgDadosUnid(UnidadeEnsino).Nome
            Call Relatorio(rptCarteiraID, strSQL)
            
            'If ChkExisteArquivo(PathBD & "\Database\IMG\" & Format(MatrID, "000000000") & "001.jpg") = True Then
            '    Set rptCarteiraID.Sections("Corpo").Controls.Item("imgFoto").Picture = LoadPicture(PathBD & "\Database\IMG\" & Format(MatrID, "000000000") & "001.jpg")
            'End If
            rptCarteiraID.Sections("Corpo").Controls.Item("lbUA").Caption = "U.A.: " & PgDadosUnid(UnidadeEnsino).UA
            
            rptCarteiraID.Show 1
            Exit Sub

            'If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            '
            'MebMatricula.PromptInclude = False
            'If MebMatricula.Text = "" Then
            '        MebMatricula.PromptInclude = True
            '        Exit Sub
            '    Else
            '        MebMatricula.PromptInclude = True
            '        If PgDadosMatr(MatrID).ValCard < Date Then
            '            MsgBox "É necessário alterar a validade da carteira!", vbInformation, "CESNet - Aviso!"
            '            Exit Sub
            '        End If
            '        If Form_Impressora.LoadFormCI(True, True, False, False, False, True, False, False, False, False) = False Then
            '            Exit Sub
            '        End If
            '        Call ImpCarteirinha(MebMatricula.Text)
            'End If

        Case Else
            Exit Sub
    End Select
End Sub

Private Sub Txt_AnoHistEsc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub



Private Sub Txt_CidadeHistEsc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub



Private Sub Txt_Comp_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub

Private Sub Txt_End_GotFocus()
    Txt_End.SelStart = 0
    Txt_End.SelLength = Len(Txt_End.Text)
End Sub
Private Sub Txt_End_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Bai_GotFocus()
    Txt_Bai.SelStart = 0
    Txt_Bai.SelLength = Len(Txt_Bai.Text)
End Sub
Private Sub Txt_Bai_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub







Private Sub Txt_EstabHistEsc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Mun_GotFocus()
    Txt_Mun.SelStart = 0
    Txt_Mun.SelLength = Len(Txt_Mun.Text)
End Sub
Private Sub Txt_Mun_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub CB_UF_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Meb_Cep_GotFocus()
    Meb_Cep.SelStart = 0
    Meb_Cep.SelLength = Len(Meb_Cep.Text) + 1
End Sub
Private Sub Meb_Cep_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
        If Acao = 1 Then PgEndereco (Meb_Cep.Text)
    End If
End Sub
Private Sub PgEndereco(CEP As String)
    'PEGAR O ENDERECO DENTRO DA TAB MATRICULA
    If Trim(CEP) = "" Then Exit Sub
    Dim RsEnd As Recordset
    
    Set RsEnd = BD.OpenRecordset("SELECT * FROM Matriculas WHERE CEP = '" & CEP & "'")
    If RsEnd.BOF And RsEnd.EOF Then
            RsEnd.Close
            Txt_End.Text = ""
            Txt_Num.Text = ""
            Txt_Comp.Text = ""
            Txt_Bai.Text = ""
            Txt_Mun.Text = ""
            CB_UF.Text = " "
        Else
            RsEnd.MoveLast
            
            Txt_End.Text = IIf(IsNull(RsEnd.Fields("End")), "", RsEnd.Fields("End"))
            Txt_Num.Text = "" 'IIf(IsNull(RsEnd.Fields("Numero")), "", RsEnd.Fields("Numero"))
            Txt_Comp.Text = "" 'IIf(IsNull(RsEnd.Fields("Compl")), "", RsEnd.Fields("Compl"))
            Txt_Bai.Text = IIf(IsNull(RsEnd.Fields("Bai")), "", RsEnd.Fields("Bai"))
            Txt_Mun.Text = IIf(IsNull(RsEnd.Fields("Mun")), "", RsEnd.Fields("Mun"))
            CB_UF.Text = IIf(IsNull(RsEnd.Fields("UF")), " ", RsEnd.Fields("UF"))
            RsEnd.Close
            Txt_Num.SetFocus
    End If
End Sub
Private Sub Cb_Sexo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Meb_Nasc_GotFocus()
    Meb_Nasc.SelStart = 0
    Meb_Nasc.SelLength = 10
End Sub
Private Sub Meb_Nasc_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Mail_GotFocus()
    Txt_Mail.SelStart = 0
    Txt_Mail.SelLength = Len(Txt_Mail.Text)
End Sub
Private Sub Txt_Mail_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(LCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Cel_GotFocus()
    Txt_Cel.SelStart = 0
    Txt_Cel.SelLength = 14
End Sub
Private Sub Txt_Cel_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Num_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub

Private Sub Txt_Tel1_GotFocus()
    Txt_Tel1.SelStart = 0
    Txt_Tel1.SelLength = 14
End Sub
Private Sub Txt_Tel1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Tel2_GotFocus()
    Txt_Tel2.SelStart = 0
    Txt_Tel2.SelLength = 14
End Sub
Private Sub Txt_Tel2_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub




Private Sub Txt_Obs_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_RG_GotFocus()
    Txt_RG.SelStart = 0
    Txt_RG.SelLength = Len(Txt_RG.Text)
End Sub
Private Sub Txt_RG_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_CertNasc_GotFocus()
    Txt_CertNasc.SelStart = 0
    Txt_CertNasc.SelLength = Len(Txt_CertNasc.Text)
End Sub
Private Sub Txt_CertNasc_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub CB_EstCiv_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_OrgEmi_GotFocus()
    Txt_OrgEmi.SelStart = 0
    Txt_OrgEmi.SelLength = Len(Txt_OrgEmi.Text)
End Sub
Private Sub Txt_OrgEmi_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
     If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Nat_GotFocus()
    Txt_Nat.SelStart = 0
    Txt_Nat.SelLength = Len(Txt_Nat.Text)
End Sub
Private Sub Txt_Nat_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Nac_GotFocus()
    Txt_Nac.SelStart = 0
    Txt_Nac.SelLength = Len(Txt_Nac.Text)
End Sub
Private Sub Txt_Nac_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Mae_GotFocus()
    Txt_Mae.SelStart = 0
    Txt_Mae.SelLength = Len(Txt_Mae.Text)
End Sub
Private Sub Txt_Mae_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_Pai_GotFocus()
    Txt_Pai.SelStart = 0
    Txt_Pai.SelLength = Len(Txt_Pai.Text)
End Sub
Private Sub Txt_Pai_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub


Private Sub LstHstEscolar()
    MSFG_HistEscolar.Rows = 1
    MSFG_HistEscolar.Rows = 2
    Set RsHistEscolar = BD.OpenRecordset("SELECT * FROM HistEscolar WHERE MatrID = '" & MatrID & "' ORDER BY Serie")
    If RsHistEscolar.BOF And RsHistEscolar.EOF Then
            'Bt_ImprCert.Enabled = False
            Exit Sub
        Else
            RsHistEscolar.MoveFirst
            lin = 1
            Do Until RsHistEscolar.EOF
                MSFG_HistEscolar.TextMatrix(lin, 0) = RsHistEscolar.Fields("Serie")
                MSFG_HistEscolar.TextMatrix(lin, 1) = IIf(IsNull(RsHistEscolar.Fields("Ano")), " ", RsHistEscolar.Fields("Ano"))
                MSFG_HistEscolar.TextMatrix(lin, 2) = IIf(IsNull(RsHistEscolar.Fields("Escola")), " ", RsHistEscolar.Fields("Escola"))
                MSFG_HistEscolar.TextMatrix(lin, 3) = IIf(IsNull(RsHistEscolar.Fields("Cidade")), " ", RsHistEscolar.Fields("Cidade"))
                RsHistEscolar.MoveNext
                MSFG_HistEscolar.Rows = MSFG_HistEscolar.Rows + 1
                lin = lin + 1
            Loop
            MSFG_HistEscolar.Rows = MSFG_HistEscolar.Rows - 1
    End If
End Sub

Private Sub txtOpcaoRel_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub



Private Sub txtRaca_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub

