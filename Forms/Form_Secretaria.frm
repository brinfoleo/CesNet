VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Secretaria 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet  - Secretaria"
   ClientHeight    =   6930
   ClientLeft      =   585
   ClientTop       =   510
   ClientWidth     =   12780
   Icon            =   "Form_Secretaria.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6930
   ScaleWidth      =   12780
   Begin VB.Frame Frame13 
      Caption         =   "Filtros:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4470
      Left            =   9900
      TabIndex        =   59
      Top             =   330
      Width           =   2835
      Begin VB.CheckBox Chk_Concluidos 
         Caption         =   "Listar Somente os Concluidos"
         Height          =   285
         Left            =   225
         TabIndex        =   112
         Top             =   225
         Value           =   1  'Checked
         Width           =   2400
      End
      Begin VB.Frame Frame16 
         Caption         =   "Ultima Ocorrencia:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   915
         Left            =   120
         TabIndex        =   71
         Top             =   3060
         Width           =   2655
         Begin VB.ComboBox Cb_FiltroOcorrencia 
            Height          =   315
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   72
            Top             =   420
            Width           =   2475
         End
      End
      Begin VB.CommandButton Bt_AplicarFiltro 
         Caption         =   "Aplicar Filtro"
         Height          =   360
         Left            =   405
         TabIndex        =   70
         Top             =   4020
         Width           =   2175
      End
      Begin VB.Frame Frame10 
         Caption         =   "Localizar Matricula:"
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
         Left            =   120
         TabIndex        =   68
         Top             =   2160
         Width           =   2655
         Begin MSMask.MaskEdBox Meb_Matricula 
            Height          =   375
            Left            =   180
            TabIndex        =   69
            Top             =   300
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
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "##.###.####"
            PromptChar      =   "_"
         End
      End
      Begin VB.ComboBox Cb_FiltroEnsino 
         Height          =   315
         ItemData        =   "Form_Secretaria.frx":030A
         Left            =   720
         List            =   "Form_Secretaria.frx":030C
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   555
         Width           =   1875
      End
      Begin VB.Frame Frame14 
         Height          =   1095
         Left            =   120
         TabIndex        =   60
         Top             =   1020
         Width           =   2655
         Begin VB.CheckBox Chk_FiltroPeriodo 
            Caption         =   "Periodo de Conclusão:"
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
            Left            =   60
            TabIndex        =   67
            Top             =   0
            Width           =   2295
         End
         Begin MSComCtl2.DTPicker DTP_FiltroPeriodoAte 
            Height          =   315
            Left            =   540
            TabIndex        =   66
            Top             =   660
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   56819713
            CurrentDate     =   38977
         End
         Begin MSComCtl2.DTPicker DTP_FiltroPeriodoDe 
            Height          =   315
            Left            =   540
            TabIndex        =   65
            Top             =   240
            Width           =   1455
            _ExtentX        =   2566
            _ExtentY        =   556
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   56819713
            CurrentDate     =   38977
         End
         Begin VB.Label Label38 
            Alignment       =   1  'Right Justify
            Caption         =   "Até:"
            Height          =   195
            Left            =   120
            TabIndex        =   64
            Top             =   780
            Width           =   315
         End
         Begin VB.Label Label37 
            Alignment       =   1  'Right Justify
            Caption         =   "De:"
            Height          =   195
            Left            =   120
            TabIndex        =   63
            Top             =   300
            Width           =   255
         End
      End
      Begin VB.Label Label34 
         Alignment       =   1  'Right Justify
         Caption         =   "Curso:"
         Height          =   195
         Left            =   120
         TabIndex        =   62
         Top             =   615
         Width           =   555
      End
   End
   Begin VB.Frame Frame8 
      Caption         =   "Impressão:"
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
      Left            =   9945
      TabIndex        =   14
      Top             =   4830
      Width           =   2835
      Begin MSComDlg.CommonDialog CD_Opcoes 
         Left            =   1140
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin MSComCtl2.DTPicker DTP_DtForms 
         Height          =   315
         Left            =   660
         TabIndex        =   19
         Top             =   960
         Width           =   1515
         _ExtentX        =   2672
         _ExtentY        =   556
         _Version        =   393216
         Format          =   56819713
         CurrentDate     =   38959
      End
      Begin VB.CommandButton Bt_ImprimirForms 
         Caption         =   "Imprimir"
         Height          =   495
         Left            =   420
         TabIndex        =   17
         Top             =   1380
         Width           =   2295
      End
      Begin VB.ComboBox Cb_Formulario 
         Height          =   315
         ItemData        =   "Form_Secretaria.frx":030E
         Left            =   180
         List            =   "Form_Secretaria.frx":0310
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   540
         Width           =   2595
      End
      Begin VB.Label Label22 
         Alignment       =   1  'Right Justify
         Caption         =   "Data:"
         Height          =   195
         Left            =   120
         TabIndex        =   18
         Top             =   1020
         Width           =   495
      End
      Begin VB.Label Label21 
         Caption         =   "Formulário:"
         Height          =   195
         Left            =   180
         TabIndex        =   15
         Top             =   300
         Width           =   855
      End
   End
   Begin TabDlg.SSTab SST_Secr 
      Height          =   4215
      Left            =   60
      TabIndex        =   3
      Top             =   2640
      Width           =   9780
      _ExtentX        =   17251
      _ExtentY        =   7435
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "Dados Pessoais"
      TabPicture(0)   =   "Form_Secretaria.frx":0312
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Disciplinas"
      TabPicture(1)   =   "Form_Secretaria.frx":032E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame9"
      Tab(1).Control(1)=   "Frame2"
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "Ocorrencias"
      TabPicture(2)   =   "Form_Secretaria.frx":034A
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame7"
      Tab(2).Control(1)=   "Frame15"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "Certificado"
      TabPicture(3)   =   "Form_Secretaria.frx":0366
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame11"
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame15 
         Caption         =   "Status:"
         Height          =   1215
         Left            =   -74820
         TabIndex        =   51
         Top             =   2880
         Width           =   9465
         Begin VB.CommandButton Bt_ExcluirOcorremcia 
            Caption         =   "Excluir Ocorremcia"
            Height          =   435
            Left            =   7410
            TabIndex        =   57
            Top             =   660
            Width           =   1845
         End
         Begin MSComCtl2.DTPicker DTP_OcorrenciaConclusao 
            Height          =   315
            Left            =   1080
            TabIndex        =   54
            Top             =   240
            Width           =   1575
            _ExtentX        =   2778
            _ExtentY        =   556
            _Version        =   393216
            Format          =   56819713
            CurrentDate     =   38976
         End
         Begin VB.CommandButton Bt_IncluirOcorrencia 
            Caption         =   "Incluir Ocorrencia"
            Height          =   435
            Left            =   7410
            TabIndex        =   53
            Top             =   180
            Width           =   1845
         End
         Begin VB.ComboBox Cb_OcorrenciaConclusao 
            Height          =   315
            Left            =   1080
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   660
            Width           =   5880
         End
         Begin VB.Label Label36 
            Alignment       =   1  'Right Justify
            Caption         =   "Ocorrencia:"
            Height          =   255
            Left            =   180
            TabIndex        =   56
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label35 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
            Height          =   255
            Left            =   480
            TabIndex        =   55
            Top             =   300
            Width           =   555
         End
      End
      Begin VB.Frame Frame11 
         Height          =   3720
         Left            =   -74820
         TabIndex        =   50
         Top             =   360
         Width           =   9345
         Begin VB.CommandButton Bt_Gravar 
            Caption         =   "Gravar"
            Height          =   465
            Left            =   6615
            TabIndex        =   122
            Top             =   720
            Width           =   2535
         End
         Begin TabDlg.SSTab SST_Cert 
            Height          =   3375
            Left            =   135
            TabIndex        =   75
            Top             =   180
            Width           =   6195
            _ExtentX        =   10927
            _ExtentY        =   5953
            _Version        =   393216
            Tab             =   2
            TabHeight       =   520
            TabCaption(0)   =   "Curso Anterior"
            TabPicture(0)   =   "Form_Secretaria.frx":0382
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame18"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "D.O."
            TabPicture(1)   =   "Form_Secretaria.frx":039E
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame12"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Observações"
            TabPicture(2)   =   "Form_Secretaria.frx":03BA
            Tab(2).ControlEnabled=   -1  'True
            Tab(2).Control(0)=   "Frame17"
            Tab(2).Control(0).Enabled=   0   'False
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame18 
               Height          =   2625
               Left            =   -74910
               TabIndex        =   113
               Top             =   315
               Width           =   6000
               Begin VB.TextBox Txt_OutrasHab 
                  Height          =   285
                  Left            =   2700
                  MaxLength       =   30
                  TabIndex        =   121
                  Top             =   2070
                  Width           =   3120
               End
               Begin VB.TextBox Txt_LocUF 
                  Height          =   285
                  Left            =   2700
                  MaxLength       =   30
                  TabIndex        =   120
                  Top             =   1530
                  Width           =   3075
               End
               Begin VB.TextBox Txt_Estab 
                  Height          =   285
                  Left            =   2700
                  MaxLength       =   30
                  TabIndex        =   119
                  Top             =   855
                  Width           =   2715
               End
               Begin VB.TextBox Txt_CursoAnt 
                  Height          =   285
                  Left            =   2700
                  MaxLength       =   20
                  TabIndex        =   118
                  Top             =   360
                  Width           =   2580
               End
               Begin VB.Label Label31 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Outras Habilitações:"
                  Height          =   195
                  Left            =   1035
                  TabIndex        =   117
                  Top             =   2070
                  Width           =   1545
               End
               Begin VB.Label Label30 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Local/UF:"
                  Height          =   195
                  Left            =   1800
                  TabIndex        =   116
                  Top             =   1530
                  Width           =   780
               End
               Begin VB.Label Label29 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Estabelecimento:"
                  Height          =   195
                  Left            =   1350
                  TabIndex        =   115
                  Top             =   945
                  Width           =   1230
               End
               Begin VB.Label Label2 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Curso Anterior e Ano de Conclusão:"
                  Height          =   195
                  Left            =   45
                  TabIndex        =   114
                  Top             =   405
                  Width           =   2535
               End
            End
            Begin VB.Frame Frame17 
               Height          =   2835
               Left            =   120
               TabIndex        =   87
               Top             =   360
               Width           =   5910
               Begin VB.TextBox Txt_ObsCert 
                  Height          =   2535
                  Left            =   120
                  MaxLength       =   50
                  MultiLine       =   -1  'True
                  ScrollBars      =   2  'Vertical
                  TabIndex        =   88
                  Top             =   180
                  Width           =   5490
               End
            End
            Begin VB.Frame Frame12 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2655
               Left            =   -74880
               TabIndex        =   76
               Top             =   465
               Width           =   4515
               Begin VB.TextBox Txt_Registro 
                  Height          =   285
                  Left            =   1500
                  MaxLength       =   5
                  TabIndex        =   81
                  Top             =   315
                  Width           =   2775
               End
               Begin VB.TextBox Txt_FolhaReg 
                  Height          =   285
                  Left            =   1485
                  MaxLength       =   10
                  TabIndex        =   80
                  Top             =   660
                  Width           =   2775
               End
               Begin VB.TextBox Txt_Livro 
                  Height          =   285
                  Left            =   1500
                  MaxLength       =   10
                  TabIndex        =   79
                  Top             =   1020
                  Width           =   2775
               End
               Begin VB.TextBox Txt_FolhaList 
                  Height          =   285
                  Left            =   1500
                  MaxLength       =   5
                  TabIndex        =   77
                  Top             =   1740
                  Width           =   2775
               End
               Begin MSComCtl2.DTPicker DTP_DtList 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   78
                  Top             =   1380
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   56819713
                  CurrentDate     =   38996
               End
               Begin VB.Label Label46 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Folha:"
                  Height          =   195
                  Left            =   975
                  TabIndex        =   86
                  Top             =   1800
                  Width           =   435
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Data:"
                  Height          =   195
                  Left            =   1005
                  TabIndex        =   85
                  Top             =   1440
                  Width           =   390
               End
               Begin VB.Label Label33 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Livro:"
                  Height          =   195
                  Left            =   1020
                  TabIndex        =   84
                  Top             =   1080
                  Width           =   375
               End
               Begin VB.Label Label32 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Folha:"
                  Height          =   195
                  Left            =   990
                  TabIndex        =   83
                  Top             =   720
                  Width           =   435
               End
               Begin VB.Label Label24 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Registro:"
                  Height          =   195
                  Left            =   825
                  TabIndex        =   82
                  Top             =   420
                  Width           =   615
               End
            End
         End
      End
      Begin VB.Frame Frame9 
         Caption         =   "DISCIPLINA concluida em:"
         Height          =   1980
         Left            =   -74880
         TabIndex        =   20
         Top             =   2025
         Width           =   9180
         Begin VB.ComboBox Cb_UFConclusao 
            Enabled         =   0   'False
            Height          =   315
            ItemData        =   "Form_Secretaria.frx":03D6
            Left            =   720
            List            =   "Form_Secretaria.frx":042B
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   1395
            Width           =   750
         End
         Begin VB.TextBox Txt_CidadeConclusao 
            Enabled         =   0   'False
            Height          =   285
            Left            =   720
            MaxLength       =   30
            TabIndex        =   23
            Top             =   1020
            Width           =   3750
         End
         Begin VB.CommandButton Bt_GrvConclusao 
            Caption         =   "Concluir Disciplina"
            Enabled         =   0   'False
            Height          =   510
            Left            =   5145
            TabIndex        =   22
            Top             =   1335
            Width           =   3825
         End
         Begin VB.ComboBox Cb_LocConclusao 
            Enabled         =   0   'False
            Height          =   315
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   630
            Width           =   3810
         End
         Begin MSMask.MaskEdBox Meb_DtConclusao 
            Height          =   315
            Left            =   705
            TabIndex        =   25
            Top             =   225
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
         Begin VB.Label Label43 
            Alignment       =   1  'Right Justify
            Caption         =   "UF:"
            Height          =   195
            Left            =   270
            TabIndex        =   29
            Top             =   1485
            Width           =   330
         End
         Begin VB.Label Label42 
            Alignment       =   1  'Right Justify
            Caption         =   "Cidade:"
            Height          =   240
            Left            =   90
            TabIndex        =   28
            Top             =   1080
            Width           =   555
         End
         Begin VB.Label Label28 
            Alignment       =   1  'Right Justify
            Caption         =   "Local:"
            Height          =   195
            Left            =   120
            TabIndex        =   27
            Top             =   735
            Width           =   435
         End
         Begin VB.Label Label25 
            Alignment       =   1  'Right Justify
            Caption         =   "Data:"
            Height          =   195
            Left            =   120
            TabIndex        =   26
            Top             =   360
            Width           =   375
         End
      End
      Begin VB.Frame Frame7 
         Height          =   2535
         Left            =   -74820
         TabIndex        =   12
         Top             =   360
         Width           =   9465
         Begin MSFlexGridLib.MSFlexGrid MSFG_Ocorrencias 
            Height          =   2295
            Left            =   120
            TabIndex        =   13
            Top             =   180
            Width           =   9180
            _ExtentX        =   16193
            _ExtentY        =   4048
            _Version        =   393216
            Cols            =   3
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"Form_Secretaria.frx":049A
         End
      End
      Begin VB.Frame Frame3 
         Height          =   3735
         Left            =   75
         TabIndex        =   6
         Top             =   360
         Width           =   9495
         Begin TabDlg.SSTab SST_DadosPessoais 
            Height          =   2775
            Left            =   90
            TabIndex        =   30
            Top             =   600
            Width           =   7320
            _ExtentX        =   12912
            _ExtentY        =   4895
            _Version        =   393216
            Tabs            =   4
            Tab             =   1
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Endereço"
            TabPicture(0)   =   "Form_Secretaria.frx":0548
            Tab(0).ControlEnabled=   0   'False
            Tab(0).Control(0)=   "Frame5"
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Documentos"
            TabPicture(1)   =   "Form_Secretaria.frx":0564
            Tab(1).ControlEnabled=   -1  'True
            Tab(1).Control(0)=   "Frame_Proprio"
            Tab(1).Control(0).Enabled=   0   'False
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Filiação"
            TabPicture(2)   =   "Form_Secretaria.frx":0580
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame6"
            Tab(2).ControlCount=   1
            TabCaption(3)   =   "Contatos"
            TabPicture(3)   =   "Form_Secretaria.frx":059C
            Tab(3).ControlEnabled=   0   'False
            Tab(3).Control(0)=   "Frame4"
            Tab(3).ControlCount=   1
            Begin VB.Frame Frame5 
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
               Height          =   2250
               Left            =   -74940
               TabIndex        =   97
               Top             =   360
               Width           =   7155
               Begin VB.ComboBox Cb_Sexo 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_Secretaria.frx":05B8
                  Left            =   2880
                  List            =   "Form_Secretaria.frx":05C5
                  Style           =   2  'Dropdown List
                  TabIndex        =   102
                  Top             =   1260
                  Width           =   855
               End
               Begin VB.ComboBox CB_UF 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_Secretaria.frx":05D2
                  Left            =   6300
                  List            =   "Form_Secretaria.frx":0627
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   101
                  Top             =   750
                  Width           =   750
               End
               Begin VB.TextBox Txt_Mun 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   3960
                  MaxLength       =   30
                  TabIndex        =   100
                  Top             =   750
                  Width           =   1815
               End
               Begin VB.TextBox Txt_Bai 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   900
                  MaxLength       =   30
                  TabIndex        =   99
                  Top             =   750
                  Width           =   1995
               End
               Begin VB.TextBox Txt_End 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   900
                  MaxLength       =   50
                  TabIndex        =   98
                  Top             =   225
                  Width           =   6135
               End
               Begin MSMask.MaskEdBox Meb_Nasc 
                  Height          =   360
                  Left            =   4980
                  TabIndex        =   103
                  Top             =   1260
                  Width           =   1215
                  _ExtentX        =   2143
                  _ExtentY        =   635
                  _Version        =   393216
                  AutoTab         =   -1  'True
                  Enabled         =   0   'False
                  MaxLength       =   10
                  Format          =   "DD/MM/YYYY"
                  Mask            =   "##/##/####"
                  PromptChar      =   "_"
               End
               Begin MSMask.MaskEdBox Meb_Cep 
                  Height          =   300
                  Left            =   900
                  TabIndex        =   104
                  Top             =   1260
                  Width           =   1035
                  _ExtentX        =   1826
                  _ExtentY        =   529
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
                  Height          =   240
                  Left            =   480
                  TabIndex        =   111
                  Top             =   1320
                  Width           =   375
               End
               Begin VB.Label Label6 
                  Alignment       =   1  'Right Justify
                  Caption         =   "UF:"
                  Height          =   195
                  Left            =   6000
                  TabIndex        =   110
                  Top             =   810
                  Width           =   255
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Municipio:"
                  Height          =   195
                  Left            =   3120
                  TabIndex        =   109
                  Top             =   810
                  Width           =   735
               End
               Begin VB.Label Label8 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Bairro:"
                  Height          =   195
                  Left            =   420
                  TabIndex        =   108
                  Top             =   810
                  Width           =   435
               End
               Begin VB.Label Label9 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Endereço:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   107
                  Top             =   300
                  Width           =   735
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Nascimento:"
                  Height          =   240
                  Left            =   4020
                  TabIndex        =   106
                  Top             =   1320
                  Width           =   915
               End
               Begin VB.Label Label12 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Sexo:"
                  Height          =   240
                  Left            =   2340
                  TabIndex        =   105
                  Top             =   1320
                  Width           =   495
               End
            End
            Begin VB.Frame Frame4 
               Caption         =   "Contatos:"
               Height          =   1500
               Left            =   -74880
               TabIndex        =   89
               Top             =   480
               Width           =   7155
               Begin VB.TextBox Txt_Tel2 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   4905
                  MaxLength       =   14
                  TabIndex        =   93
                  Top             =   630
                  Width           =   2070
               End
               Begin VB.TextBox Txt_Cel 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   765
                  MaxLength       =   14
                  TabIndex        =   92
                  Top             =   630
                  Width           =   2325
               End
               Begin VB.TextBox Txt_Tel1 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   4905
                  MaxLength       =   14
                  TabIndex        =   91
                  Top             =   240
                  Width           =   2070
               End
               Begin VB.TextBox Txt_Mail 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   675
                  MaxLength       =   30
                  TabIndex        =   90
                  Top             =   240
                  Width           =   3195
               End
               Begin VB.Label Label15 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Celular:"
                  Height          =   255
                  Left            =   120
                  TabIndex        =   96
                  Top             =   660
                  Width           =   555
               End
               Begin VB.Label Label10 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Tel.(s):"
                  Height          =   195
                  Left            =   4305
                  TabIndex        =   95
                  Top             =   420
                  Width           =   495
               End
               Begin VB.Label Label14 
                  Alignment       =   1  'Right Justify
                  Caption         =   "E-mail:"
                  Height          =   195
                  Left            =   180
                  TabIndex        =   94
                  Top             =   285
                  Width           =   450
               End
            End
            Begin VB.Frame Frame6 
               Caption         =   "Filiação:"
               Height          =   1575
               Left            =   -74820
               TabIndex        =   44
               Top             =   360
               Width           =   7005
               Begin VB.TextBox Txt_Pai 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   600
                  MaxLength       =   50
                  TabIndex        =   46
                  Top             =   600
                  Width           =   6150
               End
               Begin VB.TextBox Txt_Mae 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   600
                  MaxLength       =   50
                  TabIndex        =   45
                  Top             =   240
                  Width           =   6150
               End
               Begin VB.Label Label26 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Pai:"
                  Height          =   255
                  Left            =   240
                  TabIndex        =   48
                  Top             =   660
                  Width           =   315
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
            End
            Begin VB.Frame Frame_Proprio 
               Caption         =   "Documentos:"
               Height          =   1575
               Left            =   75
               TabIndex        =   31
               Top             =   360
               Width           =   7170
               Begin VB.TextBox Txt_Nac 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5055
                  MaxLength       =   30
                  TabIndex        =   37
                  Top             =   1125
                  Width           =   2040
               End
               Begin VB.TextBox Txt_CertNasc 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1020
                  MaxLength       =   20
                  TabIndex        =   36
                  Top             =   720
                  Width           =   1935
               End
               Begin VB.TextBox Txt_RG 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   1020
                  MaxLength       =   20
                  TabIndex        =   35
                  Top             =   240
                  Width           =   1935
               End
               Begin VB.TextBox Txt_OrgEmi 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5040
                  MaxLength       =   20
                  TabIndex        =   34
                  Top             =   240
                  Width           =   2055
               End
               Begin VB.TextBox Txt_Nat 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   5040
                  MaxLength       =   30
                  TabIndex        =   33
                  Top             =   720
                  Width           =   2055
               End
               Begin VB.ComboBox CB_EstCiv 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_Secretaria.frx":0696
                  Left            =   1020
                  List            =   "Form_Secretaria.frx":06B2
                  Style           =   2  'Dropdown List
                  TabIndex        =   32
                  Top             =   1140
                  Width           =   1935
               End
               Begin VB.Label Label20 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Nacionalidade:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   3660
                  TabIndex        =   43
                  Top             =   1200
                  Width           =   1335
               End
               Begin VB.Label Label13 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Certidão de Nascimento:"
                  Height          =   375
                  Left            =   60
                  TabIndex        =   42
                  Top             =   600
                  Width           =   915
               End
               Begin VB.Label Label16 
                  Alignment       =   1  'Right Justify
                  Caption         =   "RG:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   540
                  TabIndex        =   41
                  Top             =   300
                  Width           =   435
               End
               Begin VB.Label Label17 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Orgão Emissor:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   3480
                  TabIndex        =   40
                  Top             =   300
                  Width           =   1515
               End
               Begin VB.Label Label18 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Estado Civil:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   39
                  Top             =   1200
                  Width           =   915
               End
               Begin VB.Label Label19 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Natural da Cidade/UF:"
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H000000FF&
                  Height          =   195
                  Left            =   3060
                  TabIndex        =   38
                  Top             =   780
                  Width           =   1935
               End
            End
         End
         Begin VB.CommandButton Bt_CancelarAluno 
            Caption         =   "Cancelar"
            Height          =   285
            Left            =   7710
            TabIndex        =   11
            Top             =   1245
            Width           =   1635
         End
         Begin VB.CommandButton Bt_GravarAluno 
            Caption         =   "Gravar"
            Height          =   285
            Left            =   7710
            TabIndex        =   10
            Top             =   915
            Width           =   1635
         End
         Begin VB.CommandButton Bt_AltAluno 
            Caption         =   "Alterar"
            Height          =   285
            Left            =   7710
            TabIndex        =   9
            Top             =   615
            Width           =   1635
         End
         Begin VB.TextBox Txt_Nome 
            Enabled         =   0   'False
            Height          =   285
            Left            =   840
            MaxLength       =   50
            TabIndex        =   8
            Top             =   210
            Width           =   6630
         End
         Begin VB.Label Label23 
            Caption         =   "- Os campos em destaque são obrigatórios para a impressão do certificado."
            ForeColor       =   &H000000FF&
            Height          =   225
            Left            =   90
            TabIndex        =   49
            Top             =   3465
            Width           =   7095
         End
         Begin VB.Label Label3 
            Alignment       =   1  'Right Justify
            Caption         =   "Nome:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   195
            Left            =   180
            TabIndex        =   7
            Top             =   270
            Width           =   615
         End
      End
      Begin VB.Frame Frame2 
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1650
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   9555
         Begin MSFlexGridLib.MSFlexGrid MSFG_Disciplinas 
            Height          =   1440
            Left            =   120
            TabIndex        =   5
            Top             =   180
            Width           =   9375
            _ExtentX        =   16536
            _ExtentY        =   2540
            _Version        =   393216
            Cols            =   5
            FixedCols       =   0
            SelectionMode   =   1
            AllowUserResizing=   1
            FormatString    =   $"Form_Secretaria.frx":070D
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
         End
      End
   End
   Begin VB.Frame Frame1 
      Height          =   2250
      Left            =   60
      TabIndex        =   0
      Top             =   360
      Width           =   9750
      Begin VB.ComboBox Cb_Result 
         Height          =   315
         ItemData        =   "Form_Secretaria.frx":0794
         Left            =   1740
         List            =   "Form_Secretaria.frx":07BC
         Style           =   2  'Dropdown List
         TabIndex        =   73
         Top             =   120
         Width           =   3255
      End
      Begin VB.CheckBox Chk_SelecionarMatr 
         Caption         =   "Selecionar Matriculas"
         Height          =   195
         Left            =   6930
         TabIndex        =   58
         Top             =   180
         Width           =   1875
      End
      Begin MSFlexGridLib.MSFlexGrid MSFG_LstMatr 
         Height          =   1725
         Left            =   60
         TabIndex        =   1
         Top             =   465
         Width           =   9540
         _ExtentX        =   16828
         _ExtentY        =   3043
         _Version        =   393216
         Cols            =   7
         BackColorSel    =   16711680
         SelectionMode   =   1
         AllowUserResizing=   1
         FormatString    =   $"Form_Secretaria.frx":086E
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Resultados por Tela:"
         Height          =   195
         Left            =   90
         TabIndex        =   74
         Top             =   180
         Width           =   1515
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "SECRETARIA"
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
      Width           =   12840
   End
End
Attribute VB_Name = "Form_Secretaria"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatricula         As Recordset
Dim RsMatrEnsino        As Recordset
Dim RsMatrDisciplina    As Recordset

Dim MatrID              As String
Dim EnsinoID            As Integer
Dim DisciplinaID        As Integer

Dim OcorrenciaID        As Integer
Dim linha               As Integer

Dim DtUltimaOcorrencia  As String 'Informa a ultima ocorrencia para nao ter de gerar outra function
Private Sub GerarDadosArqDO()
    On Error GoTo T_Erro_GerarDadosArqDO
    Dim cont        As Integer
    
    Dim contSeq     As Integer
    Dim Ano         As String
    Dim strTxt      As String
    Dim NomeArq     As String
    Dim GrvOcorr    As Integer
    
    If EnsinoID = 0 Then
        MsgBox "Selecione um Ensino.", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
        
    NomeArq = "DO-" & Format(DTP_DtForms.Value, "yyyymmdd") & "-" & ConvNum(PgDadosUnid(UnidadeEnsino).Cnpj) & ".doc"

    CD_Opcoes.FileName = NomeArq
    CD_Opcoes.ShowSave
    
    If Trim(CD_Opcoes.FileName) = NomeArq Then
        Exit Sub
    End If
    NomeArq = CD_Opcoes.FileName
  
    Ano = ""
    If MsgBox("Deseja registrar uma ocorrencia para a(s) matricula(s)?", vbInformation + vbYesNo, "CESNet - Ocorrencia") = vbYes Then
            GrvOcorr = Form_SecretariaSelOcorrencia.SelOcorrencia
        Else
            GrvOcorr = 0
    End If
    '****************** Cab do Arquivo *****************************
    Call GrvArq(NomeArq, "Ensino " & Cb_FiltroEnsino.Text)
    'Call GrvArq(NomeArq, "SECRETARIA DE ESTADO DE EDUCAÇÃO")
    'Call GrvArq(NomeArq, "SUBSECRETARIA ADJUNTA DE PLANEJAMENTO PEDAGÓGICO")
    'Call GrvArq(NomeArq, "CENTRO DE ESTUDOS SUPLETIVOS" & UnidadeEnsino)
    'Call GrvArq(NomeArq, "EDITAL")
    '
    'Call GrvArq(NomeArq, "O DIRETOR DO CENTRO DE ESTUDOS SUPLETIVOS" & UnidadeEnsino & ",")
    'Call GrvArq(NomeArq, "Coordenador Regional da Região Metropolitana III, Municipio")
    '***************************************************************
    For cont = 1 To MSFG_LstMatr.Rows - 1
        If Trim(MSFG_LstMatr.TextMatrix(cont, 0)) = "X" Then
            contSeq = contSeq + 1
            'Grava o ano
            If Ano <> Right(Trim(MSFG_LstMatr.TextMatrix(cont, 3)), 4) Then
                Ano = Right(Trim(MSFG_LstMatr.TextMatrix(cont, 3)), 4)
                contSeq = "1"
                strTxt = String(4, " ") & "ANO: " & Ano
                Call GrvArq(NomeArq, strTxt)
            End If
            'Grava a sequencia de dados
            strTxt = left(String(2, "0"), 2 - Len(Trim(contSeq))) & contSeq & _
                     " - " & PgDadosMatr(MSFG_LstMatr.TextMatrix(cont, 1)).Nome
            Call GrvArq(NomeArq, strTxt)
            If GrvOcorr <> 0 Then
                If GrvUltimaOcorrenciaMatrEnsino(MSFG_LstMatr.TextMatrix(cont, 1), _
                                                 GrvOcorr, _
                                                 DTP_DtForms.Value) = True Then
                    Call GrvOcorrencia(MSFG_LstMatr.TextMatrix(cont, 1), _
                                       PgIDEnsino(MSFG_LstMatr.TextMatrix(cont, 4)), _
                                       DTP_DtForms.Value, _
                                       GrvOcorr)
                    'Call ListOcorrencias(MatrID, EnsinoID)
                End If
            End If
            
        End If
    Next
    MsgBox "Arquivo gravado com sucesso!", vbInformation, "CESNet - Aviso!"
    
    
    Exit Sub
T_Erro_GerarDadosArqDO:
    Call RegLogErros(Err.Number, Err.Description, Me.Name, UsuarioID)
    MsgBox "Erro ao gerar arquivo. Por favor tente novamente!" & vbCrLf & vbCrLf & "Descrição: " & Err.Description, vbInformation, "CESNet - Erro n.: " & Err.Number
    Exit Sub
End Sub

Private Sub HDBtAluno(op As Boolean)
    Bt_AltAluno.Enabled = op
    Bt_GravarAluno.Enabled = IIf(op = True, False, True)
    Bt_CancelarAluno.Enabled = IIf(op = True, False, True)
    
End Sub
Private Sub HDFormDados(op As Boolean)
    'MebMatricula.Enabled = False
    MSFG_LstMatr.Enabled = IIf(op = False, True, False)
    Txt_Nome.Enabled = op
    Txt_End.Enabled = op
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
    Txt_CertNasc.Enabled = op
    Txt_Nat.Enabled = op
    CB_EstCiv.Enabled = op
    Txt_Nac.Enabled = op
    
    'Cb_Deficiencia.Enabled = op
    
    Txt_Mae.Enabled = op
    Txt_Pai.Enabled = op

End Sub

Private Sub Bt_AltAluno_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub

    If Trim(MatrID) = "" Then Exit Sub
    
    HDFormDados (True)
    HDBtAluno (False)
End Sub


Private Sub Bt_AplicarFiltro_Click()
    If Trim(Cb_FiltroEnsino.Text) = "" Then
            Cb_FiltroEnsino.AddItem "(Todos)"
            Cb_FiltroEnsino.Text = "(Todos)"
        Else
            If Cb_FiltroEnsino.Text = "(Todos)" Then
                    EnsinoID = 0
                Else
                    EnsinoID = PgIDEnsino(Cb_FiltroEnsino.Text)
            End If
    End If
    Meb_Matricula.PromptInclude = False
    If Trim(Meb_Matricula.Text) = "" Then
            Meb_Matricula.PromptInclude = True
            ListMatriculas
        Else
            Meb_Matricula.PromptInclude = True
            ListMatriculas (Trim(Meb_Matricula.Text))
    End If
    SST_Secr.Tab = 0
End Sub

Private Sub Bt_CancelarAluno_Click()
    HDFormDados (False)
    HDBtAluno (True)
    
End Sub

Private Sub Bt_ExcluirOcorremcia_Click()
    If OcorrenciaID = 0 Then Exit Sub
    Dim RsEnsinoOcorrenciaConclusao As Recordset
    Set RsEnsinoOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM EnsinoOcorrenciaConclusao WHERE ContID = " & OcorrenciaID & " AND MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " ORDER BY Dt ASC")
    If RsEnsinoOcorrenciaConclusao.BOF And RsEnsinoOcorrenciaConclusao.EOF Then
            MsgBox "Nenhuma Ocorrencia encontrada.", vbInformation, "CESNet - Aviso"
        Else
            RsEnsinoOcorrenciaConclusao.MoveFirst
            RsEnsinoOcorrenciaConclusao.Delete
            Call ListOcorrencias(MatrID, EnsinoID)
    End If
    RsEnsinoOcorrenciaConclusao.Close
    
    GrvUltimaOcorrenciaMatrEnsino
    Call ListOcorrencias(MatrID, EnsinoID)
    If MSFG_Ocorrencias.Rows = 1 Then
            MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 5) = "[Nenhuma Ocorrencia Cadastrada...] Aguardando atualização"
        Else
            MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 5) = "[" & MSFG_Ocorrencias.TextMatrix(MSFG_Ocorrencias.Rows - 1, 2) & "] Aguardando atualização"
    End If
End Sub

Private Sub Bt_Gravar_Click()
    Dim RsCertificado As Recordset
    Set RsCertificado = BD.OpenRecordset("SELECT * FROM Certificado WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsCertificado.BOF And RsCertificado.EOF Then
            RsCertificado.AddNew
        Else
            RsCertificado.MoveFirst
            RsCertificado.Edit
    End If
    RsCertificado.Fields("MatrID") = MatrID
    RsCertificado.Fields("EnsinoID") = EnsinoID
    RsCertificado.Fields("CursoAnt") = IIf(Trim(Txt_CursoAnt.Text) = "", Null, Trim(Txt_CursoAnt.Text))
    RsCertificado.Fields("Estabelecimento") = IIf(Trim(Txt_Estab.Text) = "", Null, Trim(Txt_Estab.Text))
    RsCertificado.Fields("LocalUF") = IIf(Trim(Txt_LocUF.Text) = "", Null, Trim(Txt_LocUF.Text))
    RsCertificado.Fields("OutrasHab") = IIf(Trim(Txt_OutrasHab.Text) = "", Null, Trim(Txt_OutrasHab.Text))
    
    RsCertificado.Fields("Obs") = IIf(Trim(Txt_ObsCert.Text) = "", Null, Trim(Txt_ObsCert.Text))
    
    RsCertificado.Fields("DOReg") = IIf(Trim(Txt_Registro.Text) = "", Null, Trim(Txt_Registro.Text))
    RsCertificado.Fields("DOFolha") = IIf(Trim(Txt_FolhaReg.Text) = "", Null, Trim(Txt_FolhaReg.Text))
    RsCertificado.Fields("DOLivro") = IIf(Trim(Txt_Livro.Text) = "", Null, Trim(Txt_Livro.Text))
    RsCertificado.Fields("DODtPublicacao") = IIf(Trim(DTP_DtList.Value) = "", Null, Trim(DTP_DtList.Value))
    RsCertificado.Fields("DOFolhaDO") = IIf(Trim(Txt_FolhaList.Text) = "", Null, Trim(Txt_FolhaList.Text))
    RsCertificado.Update
    
    
    
End Sub

Private Sub Bt_GravarAluno_Click()
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID = '" & MatrID & "'")
    If RsMatricula.BOF And RsMatricula.EOF Then
            MsgBox "Erro ao localiza a matricula." & vbCrLf & "Operação cancelada.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsMatricula.MoveFirst
    End If
    With RsMatricula
        .Edit
        .Fields("Nome") = Trim(Txt_Nome.Text)
        .Fields("End") = IIf(Trim(Txt_End.Text) = "", Null, Trim(Txt_End.Text))
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
        .Fields("CertNasc") = IIf(Trim(Txt_CertNasc.Text) = "", Null, Trim(Txt_CertNasc.Text))
        .Fields("Natural") = IIf(Trim(Txt_Nat.Text) = "", Null, Trim(Txt_Nat.Text))
        .Fields("EstCivil") = IIf(Trim(CB_EstCiv.Text) = "", Null, Trim(CB_EstCiv.Text))
        .Fields("Nacion") = IIf(Trim(Txt_Nac.Text) = "", Null, Trim(Txt_Nac.Text))
                    
        '.Fields("DefID") = IIf(Trim(Cb_Deficiencia.Text) = "", Null, PgIDDef(Trim(Cb_Deficiencia.Text)))
                    
        .Fields("Mae") = IIf(Trim(Txt_Mae.Text) = "", Null, Trim(Txt_Mae.Text))
        .Fields("Pai") = IIf(Trim(Txt_Pai.Text) = "", Null, Trim(Txt_Pai.Text))
        .Update
        'MatrID = MebMatricula.Text
        
    End With
    Bt_CancelarAluno_Click
    MsgBox "Matricula: " & MatrID & " alterada com sucesso!", vbDefaultButton1, "CESNet - Nova Matricula"
    Call RegLog(MatrID, "Alterou os dados pessoais do aluno.")
End Sub


Private Sub Bt_ImprimirForms_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    If Chk_SelecionarMatr.Value = 0 Then
        If Trim(MatrID) = "" Then
            MsgBox "Selecione uma Matricula.", vbInformation, "CESNet - Aviso"
            Exit Sub
        End If
        If EnsinoID = 0 Then
            MsgBox "Selecione um ensino!", vbInformation, "CESNet - Aviso"
            Exit Sub
        End If
    End If
    Select Case left(Cb_Formulario.Text, 3)
        Case "001" 'HISTORICO ESCOLAR
            Call Rpt006(MatrID, EnsinoID, DTP_DtForms.Value)
            
        Case "002"
            Call Rpt005(MatrID, EnsinoID, DTP_DtForms.Value)
            
        Case "003"
            GerarDadosArqDO
        
        Case "004"
            Call Rpt002(MatrID, EnsinoID, DTP_DtForms.Value)
            
        Case "005"
            ImpCert (EnsinoID)
            
        Case Else
            MsgBox "Formulário não encontrado.", vbInformation, "CESNet - Aviso"
            Exit Sub
    End Select

End Sub

Private Sub Bt_IncluirOcorrencia_Click()
    If Trim(MatrID) = "" Then Exit Sub
    If Trim(Cb_OcorrenciaConclusao.Text) = "" Then Exit Sub
    
    If GrvUltimaOcorrenciaMatrEnsino(MatrID, Trim(left(Cb_OcorrenciaConclusao.Text, 3)), DTP_OcorrenciaConclusao.Value) = True Then
        Call GrvOcorrencia(MatrID, EnsinoID, DTP_OcorrenciaConclusao.Value, Trim(left(Cb_OcorrenciaConclusao.Text, 3)))
        MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 5) = Cb_OcorrenciaConclusao.Text
        MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 6) = DTP_OcorrenciaConclusao.Value
        Call ListOcorrencias(MatrID, EnsinoID)
    End If
    
    
End Sub
Private Sub GrvOcorrencia(nMatrID As String, nEnsinoID As Integer, nDtOcorrenciaConclusao As String, _
                          nOcorrenciaID As Integer)
    Dim RsEnsinoOcorrenciaConclusao     As Recordset
    Set RsEnsinoOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM EnsinoOcorrenciaConclusao")
    RsEnsinoOcorrenciaConclusao.AddNew
    RsEnsinoOcorrenciaConclusao.Fields("MatrID") = nMatrID
    RsEnsinoOcorrenciaConclusao.Fields("EnsinoID") = nEnsinoID
    RsEnsinoOcorrenciaConclusao.Fields("Dt") = nDtOcorrenciaConclusao
    RsEnsinoOcorrenciaConclusao.Fields("OcorrenciaID") = nOcorrenciaID
    RsEnsinoOcorrenciaConclusao.Update
    RsEnsinoOcorrenciaConclusao.Close
End Sub
Private Function GrvUltimaOcorrenciaMatrEnsino(Optional nMatrID As String, Optional Oc As Integer, Optional DtOc As String) As Boolean
    Dim RsMatriculaEnsino               As Recordset
    If Trim(nMatrID) = "" Then
        nMatrID = MatrID
    End If
    Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID ='" & nMatrID & "' AND EnsinoID = " & EnsinoID)
    If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
            MsgBox "Erro ao localizar Matricula/Ensino." & vbCrLf & "Operação cancelada!", vbInformation, "CESNet - Aviso"
            RsMatriculaEnsino.Close
            GrvUltimaOcorrenciaMatrEnsino = False
            Exit Function
        Else
            RsMatriculaEnsino.MoveFirst
            RsMatriculaEnsino.Edit
            RsMatriculaEnsino.Fields("OcorrenciaID") = IIf(PgUltimaOcorrencia(MatrID, EnsinoID) = 0, Oc, PgUltimaOcorrencia(MatrID, EnsinoID))
            RsMatriculaEnsino.Fields("DtOcorrencia") = IIf(Trim(DtUltimaOcorrencia) <> "0", DtUltimaOcorrencia, DtOc)
            RsMatriculaEnsino.Update
            RsMatriculaEnsino.Close
            GrvUltimaOcorrenciaMatrEnsino = True
       End If
End Function

Private Sub Cb_FiltroEnsino_DropDown()
    Dim RsEnsino As Recordset
    Cb_FiltroEnsino.Clear
    Cb_FiltroEnsino.AddItem "(Todos)"
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.EOF Then
            'Exit Sub
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_FiltroEnsino.AddItem RsEnsino.Fields("Descr")
                RsEnsino.MoveNext
            Loop
    End If
    RsEnsino.Close
    'Cb_FiltroEnsino.Text = "(Todos)"
End Sub






Private Sub Cb_FiltroOcorrencia_DropDown()
    'If Trim(MatrID) = "" Then Exit Sub
    Dim RsOcorrenciaConclusao As Recordset
    Cb_FiltroOcorrencia.Clear
    Set RsOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM OcorrenciaConclusao ORDER BY Descr ASC")
    If RsOcorrenciaConclusao.BOF And RsOcorrenciaConclusao.EOF Then
        Else
            RsOcorrenciaConclusao.MoveFirst
            Do Until RsOcorrenciaConclusao.EOF
                Cb_FiltroOcorrencia.AddItem left("000", 3 - Len(Trim(RsOcorrenciaConclusao.Fields("OcorrenciaID")))) & RsOcorrenciaConclusao.Fields("OcorrenciaID") & " - " & RsOcorrenciaConclusao.Fields("Descr")
                RsOcorrenciaConclusao.MoveNext
            Loop
            
    End If
    RsOcorrenciaConclusao.Close
End Sub





Private Sub Cb_OcorrenciaConclusao_DropDown()
    If Trim(MatrID) = "" Then Exit Sub
    Dim RsOcorrenciaConclusao As Recordset
    Cb_OcorrenciaConclusao.Clear
    Set RsOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM OcorrenciaConclusao ORDER BY Descr ASC")
    If RsOcorrenciaConclusao.BOF And RsOcorrenciaConclusao.EOF Then
        Else
            RsOcorrenciaConclusao.MoveFirst
            Do Until RsOcorrenciaConclusao.EOF
                Cb_OcorrenciaConclusao.AddItem left("000", 3 - Len(Trim(RsOcorrenciaConclusao.Fields("OcorrenciaID")))) & RsOcorrenciaConclusao.Fields("OcorrenciaID") & " - " & RsOcorrenciaConclusao.Fields("Descr")
                RsOcorrenciaConclusao.MoveNext
            Loop
            
    End If
    RsOcorrenciaConclusao.Close
End Sub







Private Sub Chk_FiltroPeriodo_Click()
    If Chk_FiltroPeriodo.Value = 1 Then
            DTP_FiltroPeriodoDe.Value = Date - 30
            DTP_FiltroPeriodoAte.Value = Date
            DTP_FiltroPeriodoDe.Enabled = True
            DTP_FiltroPeriodoAte.Enabled = True
        Else
            DTP_FiltroPeriodoDe.Enabled = False
            DTP_FiltroPeriodoAte.Enabled = False
    End If
End Sub

Private Sub Chk_SelecionarMatr_Click()
    If Chk_SelecionarMatr.Value = 1 Then
            SST_Secr.Enabled = False
        Else
            SST_Secr.Enabled = True
    End If
    ListMatriculas
End Sub



Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    DoEvents
    Me.top = 0
    Me.left = 0
    Me.MousePointer = 11
    LimpForm
    Cb_Result.Text = Cb_Result.List(1) '."05 resultados"
    ListMatriculas
    
    'Carrega o combo impressao
    Cb_Formulario.AddItem ("001 - Histórico escolar.")
    Cb_Formulario.AddItem ("002 - Listagem das provas efetuadas.")
    Cb_Formulario.AddItem ("003 - Gerar Listagem para o D.O.")
    Cb_Formulario.AddItem ("004 - Declaração de Conclusão.")
    Cb_Formulario.AddItem ("005 - Certificado.")
    
    '*******************************************
    Cb_FiltroEnsino.AddItem "(Todos)"
    Cb_FiltroEnsino.Text = "(Todos)"


    DTP_DtForms.Value = Date
    DTP_OcorrenciaConclusao.Value = Date
    SST_Secr.Tab = 0
    SST_DadosPessoais.Tab = 0
    
    Me.MousePointer = 0
End Sub
Private Sub LimpForm()
    MSFG_LstMatr.Rows = 1
End Sub
Private Sub ListMatriculas(Optional tmpMatr As String)
    Dim strSQL          As String
    Dim TmpOcorrencia   As String
    Dim QtdResult       As String
    linha = 0
    Me.MousePointer = 11
    If Trim(Cb_Result.Text) = "(Todos)" Or Trim(Cb_Result.Text) = "" Then
            QtdResult = "*"
        Else
            QtdResult = "TOP " & left(Trim(Cb_Result.Text), 2) & " *"
            '"SELECT TOP " & RPMax & " * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' ORDER BY DtAvaliacao DESC")
    End If
    
    strSQL = "SELECT " & QtdResult & " FROM MatriculaEnsino"
    
    'Nao listar os trancados
    strSQL = strSQL & " WHERE Trancado = FALSE"
    
    'Listar tabem os nao concluidos
    If Chk_Concluidos.Value = 1 Then
            strSQL = strSQL & " AND DtFinal <> Null"
    End If
    
    
    
    If Trim(tmpMatr) <> "" Then
            strSQL = strSQL & IIf(Right(strSQL, 6) = "Ensino", " WHERE", " AND") & " MatrID = '" & tmpMatr & "'" ' AND DtFinal <> Null ORDER BY DtFinal DESC, EnsinoID ASC"
    End If
    
    If PgIDEnsino(Cb_FiltroEnsino.Text) <> 0 Then
        strSQL = strSQL & IIf(Right(strSQL, 6) = "Ensino", " WHERE", " AND") & " EnsinoID = " & PgIDEnsino(Cb_FiltroEnsino.Text)
    End If
    
    If Chk_FiltroPeriodo.Value = 1 Then
        strSQL = strSQL & IIf(Right(strSQL, 6) = "Ensino", " WHERE", " AND") & " DtFinal BETWEEN #" & Format(DTP_FiltroPeriodoDe.Value, "mm/dd/yyyy") & "# AND #" & Format(DTP_FiltroPeriodoAte.Value, "mm/dd/yyyy") & "#"
    End If
    If IsNumeric(Trim(left(Cb_FiltroOcorrencia.Text, 3))) Then
        strSQL = strSQL & IIf(Right(strSQL, 6) = "Ensino", " WHERE", " AND") & " OcorrenciaID = " & Trim(left(Cb_FiltroOcorrencia.Text, 3))
    End If
    
    strSQL = strSQL & " ORDER BY DtFinal ASC, EnsinoID ASC"
    
    DoEvents
    MSFG_LstMatr.Rows = 1
    Set RsMatrEnsino = BD.OpenRecordset(strSQL)
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
            'MSFG_LstMatr.Rows = 1
        Else
            RsMatrEnsino.MoveFirst
            Do Until RsMatrEnsino.EOF
                With MSFG_LstMatr
                    .Rows = .Rows + 1
                    '.TextMatrix(.Rows - 1, 0) = ""
                    .TextMatrix(.Rows - 1, 1) = RsMatrEnsino.Fields("MatrID")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsMatrEnsino.Fields("DtInicio")), "SEM DATA", RsMatrEnsino.Fields("DtInicio"))
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(RsMatrEnsino.Fields("DtFinal")), " ", RsMatrEnsino.Fields("DtFinal"))
                    .TextMatrix(.Rows - 1, 4) = PgNomeEnsino(RsMatrEnsino.Fields("EnsinoID"))
                    TmpOcorrencia = IIf(IsNull(RsMatrEnsino.Fields("OcorrenciaID")), 0, RsMatrEnsino.Fields("OcorrenciaID")) 'PgUltimaOcorrencia(RsMatrEnsino.Fields("MatrID"), RsMatrEnsino.Fields("EnsinoID"))
                    If TmpOcorrencia = 0 Then
                            TmpOcorrencia = "Nenhuma Ocorrencia Cadastrada..."
                        Else
                            TmpOcorrencia = left("000", 3 - Len(Trim(TmpOcorrencia))) & TmpOcorrencia & " - " & PgNomeOcorrenciaConclusao(TmpOcorrencia)
                    End If
                    .TextMatrix(.Rows - 1, 5) = TmpOcorrencia
                    .TextMatrix(.Rows - 1, 6) = IIf(IsNull(RsMatrEnsino.Fields("DtOcorrencia")), "SEM DATA", RsMatrEnsino.Fields("DtOcorrencia"))
                    '.TextMatrix(.Rows - 1, 5) = "Não" 'RsMatrEnsino.Fields("DtFinal")
                End With
                RsMatrEnsino.MoveNext
            Loop
    End If
    Me.MousePointer = 0
End Sub





Private Sub Meb_Matricula_GotFocus()
    Meb_Matricula.SelStart = 0
    Meb_Matricula.SelLength = 11
End Sub


Private Sub Meb_Matricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then Bt_AplicarFiltro_Click
    
End Sub

Private Sub MSFG_Disciplinas_Click()
    With MSFG_Disciplinas
        If Trim(.TextMatrix(.Row, 0)) = "" Or Trim(.TextMatrix(.Row, 0)) = "Disciplina" Then
            HDDiscConcluidas (False)
            Exit Sub
        End If
        DisciplinaID = PgIDDisciplina(.TextMatrix(.Row, 0))
        Meb_DtConclusao.Text = .TextMatrix(.Row, 1)
        Cb_LocConclusao.AddItem .TextMatrix(.Row, 2)
        Cb_LocConclusao.Text = .TextMatrix(.Row, 2)
        Txt_CidadeConclusao.Text = .TextMatrix(.Row, 3)
        Cb_UFConclusao.Text = .TextMatrix(.Row, 4)
        HDDiscConcluidas (True)
    End With
    
End Sub
Private Sub HDDiscConcluidas(op As Boolean)
    Meb_DtConclusao.Enabled = op
    Cb_LocConclusao.Enabled = op
    Txt_CidadeConclusao.Enabled = op
    Cb_UFConclusao.Enabled = op
    
    Bt_GrvConclusao.Enabled = op
    
End Sub
Private Sub Cb_LocConclusao_DropDown()
    Dim RsInstEnsino    As Recordset
    Cb_LocConclusao.Clear
    Set RsInstEnsino = BD.OpenRecordset("Select * FROM InstEnsino ORDER BY Descr")
    If RsInstEnsino.BOF And RsInstEnsino.EOF Then
            Cb_LocConclusao.Clear
        Else
            RsInstEnsino.MoveFirst
            Do Until RsInstEnsino.EOF
                Cb_LocConclusao.AddItem (left("000", 3 - Len(RsInstEnsino.Fields("ID"))) & Trim(RsInstEnsino.Fields("ID")) & " - " & RsInstEnsino.Fields("Descr"))
                RsInstEnsino.MoveNext
            Loop
    End If
End Sub

Private Sub MSFG_LstMatr_Click()
    If Trim(MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 1)) = "" Or Trim(MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 1)) = "Matricula" Then Exit Sub
    
    Dim Sinal As String
    'Sinal = ">"
        
    If Chk_SelecionarMatr.Value = 1 Then
            Sinal = "X"
            If MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0) = Sinal Then
                    MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0) = ""
                Else
                    MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0) = Sinal
                    'Linha = MSFG_LstMatr.Row
            End If
        Else
            Sinal = ">"
            If linha = 0 Then
                    MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0) = Sinal
                    linha = MSFG_LstMatr.Row
                Else
                    MSFG_LstMatr.TextMatrix(linha, 0) = ""
                    MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0) = Sinal
                    linha = MSFG_LstMatr.Row
            End If
            MatrID = MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 1)
            EnsinoID = PgIDEnsino(MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 4))
            ListDadosPessoais (MatrID)
            ListDiscipl (MatrID)
            Call ListOcorrencias(MatrID, EnsinoID)

    End If
    
End Sub
Private Sub ListDadosPessoais(nMatr As String)
    Dim RsCertificado As Recordset
    
    SST_Secr.Tab = 0
    SST_DadosPessoais.Tab = 0
    Txt_Nome.Text = Trim(PgDadosMatr(nMatr).Nome) 'MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0) & " - " & PgDadosMatr(MSFG_LstMatr.TextMatrix(MSFG_LstMatr.Row, 0)).Nome
    Txt_End.Text = Trim(PgDadosMatr(nMatr).Endereco)
    Txt_Bai.Text = PgDadosMatr(nMatr).Bairro
    Txt_Mun.Text = PgDadosMatr(nMatr).Munic
    CB_UF.Text = PgDadosMatr(nMatr).UF
    Meb_Cep.Text = PgDadosMatr(nMatr).CEP
    Cb_Sexo.Text = PgDadosMatr(nMatr).Sexo
    Meb_Nasc.PromptInclude = False
    Meb_Nasc.Text = PgDadosMatr(nMatr).Nasc
    Meb_Nasc.PromptInclude = True
    Txt_Mail.Text = PgDadosMatr(nMatr).Mail
    Txt_Cel.Text = PgDadosMatr(nMatr).Cel
    Txt_Tel1.Text = PgDadosMatr(nMatr).Tel1
    Txt_Tel2.Text = PgDadosMatr(nMatr).Tel2
    Txt_RG.Text = PgDadosMatr(nMatr).RG
    Txt_OrgEmi.Text = PgDadosMatr(nMatr).OE
    Txt_CertNasc.Text = PgDadosMatr(nMatr).CertNasc
    Txt_Nat.Text = PgDadosMatr(nMatr).Natural
    CB_EstCiv.Text = PgDadosMatr(nMatr).EstCivil
    Txt_Nac.Text = PgDadosMatr(nMatr).Nacion
    Txt_Mae.Text = PgDadosMatr(nMatr).Mae
    Txt_Pai.Text = PgDadosMatr(nMatr).Pai
    
    
    
    
    'Certificado
    Set RsCertificado = BD.OpenRecordset("SELECT * FROM Certificado WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsCertificado.BOF And RsCertificado.EOF Then
            RsCertificado.Close
            Txt_CursoAnt.Text = ""
            Txt_Estab.Text = ""
            Txt_LocUF.Text = ""
            Txt_OutrasHab.Text = ""
    
            Txt_ObsCert.Text = ""
    
            Txt_Registro.Text = ""
            Txt_FolhaReg.Text = ""
            Txt_Livro.Text = ""
            DTP_DtList.Value = Date
            Txt_FolhaList.Text = ""
            Exit Sub
        Else
            RsCertificado.MoveFirst
            'RsCertificado.Fields("MatrID") = MatrID
            'RsCertificado.Fields("EnsinoID") = EnsinoID
            Txt_CursoAnt.Text = IIf(IsNull(RsCertificado.Fields("CursoAnt")), "", RsCertificado.Fields("CursoAnt"))
            Txt_Estab.Text = IIf(IsNull(RsCertificado.Fields("Estabelecimento")), "", RsCertificado.Fields("Estabelecimento"))
            Txt_LocUF.Text = IIf(IsNull(RsCertificado.Fields("LocalUF")), "", RsCertificado.Fields("LocalUF"))
            Txt_OutrasHab.Text = IIf(IsNull(RsCertificado.Fields("OutrasHab")), "", RsCertificado.Fields("OutrasHab"))
    
            Txt_ObsCert.Text = IIf(IsNull(RsCertificado.Fields("Obs")), "", RsCertificado.Fields("Obs"))
    
            Txt_Registro.Text = IIf(IsNull(RsCertificado.Fields("DOReg")), "", RsCertificado.Fields("DOReg"))
            Txt_FolhaReg.Text = IIf(IsNull(RsCertificado.Fields("DOFolha")), "", RsCertificado.Fields("DOFolha"))
            Txt_Livro.Text = IIf(IsNull(RsCertificado.Fields("DOLivro")), "", RsCertificado.Fields("DOLivro"))
            DTP_DtList.Value = IIf(IsNull(RsCertificado.Fields("DODtPublicacao")), "", RsCertificado.Fields("DODtPublicacao"))
            Txt_FolhaList.Text = IIf(IsNull(RsCertificado.Fields("DOFolhaDO")), "", RsCertificado.Fields("DOFolhaDO"))
    
    
            RsCertificado.Close
    End If
    'Txt_UnidFederacao.Text = PgNomeUF(PgDadosMatr(nMatr).UF)
    'Cb_Unidade.AddItem PgDadosMatr(nMatr).UnidMatr
    'Cb_Unidade.Text = PgDadosMatr(nMatr).UnidMatr
    'Txt_EndUnidEnsino.Text = PgDadosUnid.Endereco
    'Txt_Criacao.Text = PgDadosUnid(Left(PgDadosMatr(nMatr).UnidMatr, 3)).AtoCriacao
    'Txt_AutoCurso.Text = PgDadosUnid(Left(PgDadosMatr(nMatr).UnidMatr, 3)).AutorCurso
    'Txt_Local.Text = PgDadosUnid(Left(PgDadosMatr(nMatr).UnidMatr, 3)).Municipio

    
End Sub
Private Sub ListDiscipl(nMatr As String)
    MSFG_Disciplinas.Rows = 1
    Meb_DtConclusao.PromptInclude = False
    Meb_DtConclusao.Text = ""
    Meb_DtConclusao.PromptInclude = True
    Cb_LocConclusao.Clear
    Txt_CidadeConclusao.Text = ""
    Cb_UFConclusao.Text = " "
    
    
    
    
    Set RsMatrDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & nMatr & "' AND EnsinoID = " & EnsinoID & " ORDER BY DtConclusao ASC")
    If RsMatrDisciplina.BOF And RsMatrDisciplina.EOF Then
        Else
            RsMatrDisciplina.MoveFirst
            Do Until RsMatrDisciplina.EOF
                With MSFG_Disciplinas
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = PgNomeDisciplina(RsMatrDisciplina.Fields("DisciplinaID"))
                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(RsMatrDisciplina.Fields("DtConclusao")), "SEM DATA", RsMatrDisciplina.Fields("DtConclusao"))
                    
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsMatrDisciplina.Fields("Local")), " ", RsMatrDisciplina.Fields("Local"))
                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(RsMatrDisciplina.Fields("Cidade")), " ", RsMatrDisciplina.Fields("Cidade"))
                    .TextMatrix(.Rows - 1, 4) = IIf(IsNull(RsMatrDisciplina.Fields("UF")), " ", RsMatrDisciplina.Fields("UF"))
                End With
                RsMatrDisciplina.MoveNext
            Loop
            
    End If
End Sub
Private Sub ListOcorrencias(nMatrID As String, nEnsinoID As Integer)
    
    Dim RsEnsinoOcorrenciaConclusao As Recordset
    OcorrenciaID = 0
    MSFG_Ocorrencias.Rows = 1
    Set RsEnsinoOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM EnsinoOcorrenciaConclusao WHERE MatrID = '" & nMatrID & "' AND EnsinoID = " & nEnsinoID & " ORDER BY Dt ASC")
    If RsEnsinoOcorrenciaConclusao.BOF And RsEnsinoOcorrenciaConclusao.EOF Then
        Else
            RsEnsinoOcorrenciaConclusao.MoveFirst
            MSFG_Ocorrencias.Rows = 1
            Do Until RsEnsinoOcorrenciaConclusao.EOF
                With MSFG_Ocorrencias
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = left("000", 3 - Len(Trim(RsEnsinoOcorrenciaConclusao.Fields("ContID")))) & Trim(RsEnsinoOcorrenciaConclusao.Fields("ContID"))
                    .TextMatrix(.Rows - 1, 1) = RsEnsinoOcorrenciaConclusao.Fields("Dt")
                    .TextMatrix(.Rows - 1, 2) = left("000", 3 - Len(Trim(RsEnsinoOcorrenciaConclusao.Fields("OcorrenciaID")))) & Trim(RsEnsinoOcorrenciaConclusao.Fields("OcorrenciaID")) & " - " & _
                                                PgNomeOcorrenciaConclusao(RsEnsinoOcorrenciaConclusao.Fields("OcorrenciaID"))
                
                End With
            
                RsEnsinoOcorrenciaConclusao.MoveNext
            Loop
    End If
    RsEnsinoOcorrenciaConclusao.Close
End Sub
Private Function PgUltimaOcorrencia(nMatrID As String, nEnsinoID As Integer) As Integer
    Dim RsEnsinoOcorrenciaConclusao As Recordset
    Set RsEnsinoOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM EnsinoOcorrenciaConclusao WHERE MatrID = '" & nMatrID & "' AND EnsinoID = " & nEnsinoID & " ORDER BY Dt DESC")
    If RsEnsinoOcorrenciaConclusao.BOF And RsEnsinoOcorrenciaConclusao.EOF Then
            PgUltimaOcorrencia = 0
            DtUltimaOcorrencia = 0
        Else
            RsEnsinoOcorrenciaConclusao.MoveFirst
            PgUltimaOcorrencia = Trim(RsEnsinoOcorrenciaConclusao.Fields("OcorrenciaID"))
            DtUltimaOcorrencia = Trim(IIf(IsNull(RsEnsinoOcorrenciaConclusao.Fields("Dt")), "0", RsEnsinoOcorrenciaConclusao.Fields("Dt")))
    End If
    RsEnsinoOcorrenciaConclusao.Close
End Function

Private Sub Bt_GrvConclusao_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    
    Dim RsMatriculaDisciplina   As Recordset
    Dim RsMatriculaEnsino       As Recordset

    Meb_DtConclusao.Text = Trim(Meb_DtConclusao.Text)
    Cb_LocConclusao.Text = Trim(Cb_LocConclusao.Text)
    
    If IsNumeric(Trim(left(Cb_LocConclusao.Text, 3))) Then
        Else
            MsgBox "Favor selecionar uma instituição de ensino", vbInformation, "CESNet - Aviso"
            Exit Sub
    End If
    If Trim(Meb_DtConclusao.Text) = "" Then
        MsgBox "Favor informar a DATA da conclusão!", vbInformation, "CESNet - Aviso!"
        Meb_DtConclusao.SetFocus
        Exit Sub
    End If
    If Not IsDate(Meb_DtConclusao.Text) Then
        MsgBox "Formato MES/ANO invalido. Por favor verifique.", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    If Trim(Cb_LocConclusao.Text) = "" Then
        MsgBox "Favor informar o LOCAL da conclusão!", vbInformation, "CESNet - Aviso!"
        Cb_LocConclusao.SetFocus
        Exit Sub
    End If
    If Trim(Txt_CidadeConclusao.Text) = "" Then
        MsgBox "Favor informar a CIDADE da conclusão!", vbInformation, "CESNet - Aviso!"
        Txt_CidadeConclusao.SetFocus
        Exit Sub
    End If
    If Trim(Cb_UFConclusao.Text) = "" Then
        MsgBox "Favor informar a UF da conclusão!", vbInformation, "CESNet - Aviso!"
        Cb_UFConclusao.SetFocus
        Exit Sub
    End If
    'Checa se existe Ensino iniciado
    'Set RsTmp = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND IsNull(DtFinal)")
    'If RsTmp.BOF And RsTmp.EOF Then
            'RsTmp.AddNew
            'RsTmp.Fields("MatrID") = MatrID
            'RsTmp.Fields("EnsinoID") = EnsinoID
            'RsTmp.Update
       ' Else
            'Checa se existe ensino aberto que nao seja o editado pelo usuario
            'RsTmp.MoveFirst
            'Do Until RsTmp.EOF
                'If EnsinoID = RsTmp.Fields("EnsinoID") Then
                    'Else
                        'MsgBox "Esta matricula já possui o ensino " & PgNomeEnsino(RsTmp.Fields("EnsinoID")) & " em andamento.", vbInformation, "CESNet - Atenção"
                        'Call LimpConclDiscipl
                        'Exit Sub
                'End If
                'RsTmp.MoveNext
            'Loop
    'End If
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            MsgBox "Erro ao localizar a disciplina, tente novamente." & vbCrLf & "Operação cancelada.", vbInformation, "CESNet - Aviso"
            Exit Sub
            'RsMatriculaDisciplina.AddNew
            'RsMatriculaDisciplina.Fields("MatrID") = MatrID
            'RsMatriculaDisciplina.Fields("EnsinoID") = EnsinoID
            'RsMatriculaDisciplina.Fields("DisciplinaID") = DisciplinaID
            'Meb_DtConclusao.PromptInclude = False
            'If Trim(Meb_DtConclusao.Text) = "" Then
            '    Else
            '        Meb_DtConclusao.PromptInclude = True
             '       RsMatriculaDisciplina.Fields("DtConclusao") = Meb_DtConclusao.Text
            'End If
            'Meb_DtConclusao.PromptInclude = True
            'RsMatriculaDisciplina.Fields("Local") = IIf(Cb_LocConclusao.Text = "", Null, Cb_LocConclusao.Text)
            'RsMatriculaDisciplina.Fields("Cidade") = Trim(Txt_CidadeConclusao.Text)
            'RsMatriculaDisciplina.Fields("UF") = Trim(Cb_UFConclusao.Text)
            'RsMatriculaDisciplina.Update
            'MsgBox "Disciplina concluida com sucesso.", vbInformation, "CESNet - Atenção"
        Else
            'If Meb_DtConclusao.Text = "" And Cb_LocConclusao.Text = "" Then
                    'RsMatricula.Delete
                'Else
                    RsMatriculaDisciplina.MoveFirst
                    RsMatriculaDisciplina.Edit
                    RsMatriculaDisciplina.Fields("DtConclusao") = IIf(Meb_DtConclusao.Text = "", Null, Meb_DtConclusao.Text)
                    RsMatriculaDisciplina.Fields("Local") = IIf(Cb_LocConclusao.Text = "", Null, Trim(Mid(Cb_LocConclusao.Text, 6, Len(Cb_LocConclusao.Text))))
                    RsMatriculaDisciplina.Fields("InstID") = Trim(left(Cb_LocConclusao.Text, 3))
                    RsMatriculaDisciplina.Fields("Abrev") = PgDadosInstEns(Trim(left(Cb_LocConclusao.Text, 3))).Abreviatura
                    RsMatriculaDisciplina.Fields("Cidade") = Trim(Txt_CidadeConclusao.Text)
                    RsMatriculaDisciplina.Fields("UF") = Trim(Cb_UFConclusao.Text)
                    RsMatriculaDisciplina.Update
                    MsgBox "Disciplina alterada e concluida com sucesso.", vbInformation, "CESNet - Aviso"
            'End If
    End If
    'Checa data de conclusao do ensino
    If Chk_ConcEnsino(MatrID, EnsinoID) = True Then
        'Call Grv_ConcEnsino(MatrID, EnsinoID)
        Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " ORDER BY DtConclusao")
            If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                MsgBox "Erro ao localizar a disciplina concluida. Ensino nao concluido!", vbInformation, "CESNet - Aviso"
                Exit Sub
            End If
        Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
        If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                RsMatriculaEnsino.AddNew
                RsMatriculaEnsino.Fields("MatrID") = MatrID
                RsMatriculaEnsino.Fields("EnsinoID") = EnsinoID
                RsMatriculaDisciplina.MoveFirst
                RsMatriculaEnsino.Fields("DtInicio") = RsMatriculaDisciplina.Fields("DtConclusao")
                RsMatriculaDisciplina.MoveLast
                RsMatriculaEnsino.Fields("DtFinal") = RsMatriculaDisciplina.Fields("DtConclusao")
                RsMatriculaEnsino.Fields("Local") = RsMatriculaDisciplina.Fields("Local")
                RsMatriculaEnsino.Update
            Else
                RsMatriculaDisciplina.MoveLast
                'RsMatriculaEnsino.MoveFirst
                RsMatriculaEnsino.Edit
                RsMatriculaEnsino.Fields("DtFinal") = RsMatriculaDisciplina.Fields("DtConclusao")
                RsMatriculaEnsino.Fields("Local") = RsMatriculaDisciplina.Fields("Local")
                RsMatriculaEnsino.Update
        End If
    End If
    ListDiscipl (MatrID)
End Sub
Private Sub MSFG_Ocorrencias_Click()
    DTP_OcorrenciaConclusao.Value = Date
    Cb_OcorrenciaConclusao.Clear
    If MSFG_Ocorrencias.Rows = 1 Or Trim(MSFG_Ocorrencias.TextMatrix(MSFG_Ocorrencias.Row, 0)) = "" Then Exit Sub
    OcorrenciaID = MSFG_Ocorrencias.TextMatrix(MSFG_Ocorrencias.Row, 0)
    DTP_OcorrenciaConclusao.Value = MSFG_Ocorrencias.TextMatrix(MSFG_Ocorrencias.Row, 1)
    Cb_OcorrenciaConclusao.AddItem MSFG_Ocorrencias.TextMatrix(MSFG_Ocorrencias.Row, 2)
    Cb_OcorrenciaConclusao.Text = MSFG_Ocorrencias.TextMatrix(MSFG_Ocorrencias.Row, 2)
    
End Sub

Private Sub SST_Secr_Click(PreviousTab As Integer)
    If SST_Secr.Tab = 0 Then
        SST_DadosPessoais.Tab = 0
    End If
End Sub


Private Sub Txt_CidadeConclusao_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub


Private Sub GrvArq(LocalArquivo As String, txt As String)
    Dim fso As New FileSystemObject
    Dim Arquivo As File
    Dim arquivoLog As TextStream
    'Dim msg As String
    'Dim caminho As String
    'caminho = App.Path & "\ErrLog.txt"
'se o arquivo não existir então cria
    If fso.FileExists(LocalArquivo) Then
            Set Arquivo = fso.GetFile(LocalArquivo)
        Else
            Set arquivoLog = fso.CreateTextFile(LocalArquivo)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(LocalArquivo)
    End If
'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
'monta informações para gerar a linha com erro
    'msg = "[" & Now & "]" & Form & ":[" & Num & "-" & Descr & "]"
' inclui linhas no arquivo texto
    arquivoLog.WriteLine txt
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)
'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing

End Sub
Private Function ConvNum(str As String) As String
    Dim i       As Integer
    Dim RspStr    As String
    Dim TmpStr  As String
    For i = 1 To Len(str)
        TmpStr = Mid(str, i, 1)
        If IsNumeric(TmpStr) Then
            RspStr = RspStr & TmpStr
        End If
    Next
    ConvNum = RspStr
End Function
Private Sub ImpCert(nModelo As Integer)
    Call Form_ImpressaoCertificado.ImprimirCertificado(nModelo, MatrID, DTP_DtForms.Value)
End Sub

Private Sub Txt_CursoAnt_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_Estab_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_FolhaList_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_FolhaReg_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_Livro_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_LocUF_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_ObsCert_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_OutrasHab_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


Private Sub Txt_Registro_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


