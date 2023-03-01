VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "msflxgrd.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_Matricula 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Matricula"
   ClientHeight    =   6600
   ClientLeft      =   450
   ClientTop       =   465
   ClientWidth     =   10485
   Icon            =   "Form_Matricula.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6600
   ScaleWidth      =   10485
   Begin VB.ComboBox Cb_Nome 
      Height          =   315
      ItemData        =   "Form_Matricula.frx":030A
      Left            =   765
      List            =   "Form_Matricula.frx":030C
      Sorted          =   -1  'True
      TabIndex        =   1
      Top             =   390
      Width           =   6675
   End
   Begin MSMask.MaskEdBox MebMatricula 
      Height          =   375
      Left            =   8730
      TabIndex        =   4
      Top             =   345
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
   Begin TabDlg.SSTab SST_Matricula 
      Height          =   5235
      Left            =   180
      TabIndex        =   8
      Top             =   1200
      Width           =   10185
      _ExtentX        =   17965
      _ExtentY        =   9234
      _Version        =   393216
      TabOrientation  =   1
      Style           =   1
      Tabs            =   4
      Tab             =   3
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "Form_Matricula.frx":030E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame_Geral"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Cadastro"
      TabPicture(1)   =   "Form_Matricula.frx":032A
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Frame_Matricula"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "Disciplinas"
      TabPicture(2)   =   "Form_Matricula.frx":0346
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame_Disciplinas"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "Provas"
      TabPicture(3)   =   "Form_Matricula.frx":0362
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "Frame_Provas"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      Begin VB.Frame Frame_Provas 
         Height          =   4695
         Left            =   90
         TabIndex        =   54
         Top             =   45
         Width           =   9945
         Begin VB.ComboBox Cb_DisciplinaProvas 
            Height          =   315
            Left            =   5895
            Style           =   2  'Dropdown List
            TabIndex        =   58
            Top             =   540
            Width           =   3795
         End
         Begin VB.ComboBox Cb_EnsinoProvas 
            Height          =   315
            Left            =   810
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   540
            Width           =   3615
         End
         Begin VB.Frame Frame10 
            Height          =   3690
            Left            =   135
            TabIndex        =   55
            Top             =   900
            Width           =   9705
            Begin MSFlexGridLib.MSFlexGrid MSFG_HstProvas 
               Height          =   3435
               Left            =   60
               TabIndex        =   56
               Top             =   180
               Width           =   9555
               _ExtentX        =   16854
               _ExtentY        =   6059
               _Version        =   393216
               Cols            =   8
               FixedCols       =   0
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   $"Form_Matricula.frx":037E
            End
         End
         Begin VB.Label Label30 
            Caption         =   "Curso:"
            Height          =   255
            Left            =   225
            TabIndex        =   61
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label29 
            Caption         =   "Disciplina:"
            Height          =   195
            Left            =   5130
            TabIndex        =   60
            Top             =   585
            Width           =   720
         End
         Begin VB.Label Label24 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "P R O V A S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   59
            Top             =   135
            Width           =   9825
         End
      End
      Begin VB.Frame Frame_Disciplinas 
         Height          =   4695
         Left            =   -74910
         TabIndex        =   48
         Top             =   45
         Width           =   9825
         Begin VB.CommandButton btExcluirDisciplina 
            Enabled         =   0   'False
            Height          =   495
            Left            =   9120
            Picture         =   "Form_Matricula.frx":0484
            Style           =   1  'Graphical
            TabIndex        =   67
            ToolTipText     =   "Excluir disciplina..."
            Top             =   420
            Width           =   555
         End
         Begin VB.Frame Frame9 
            Height          =   3630
            Left            =   120
            TabIndex        =   50
            Top             =   960
            Width           =   9585
            Begin MSFlexGridLib.MSFlexGrid MSFG_HstDisciplinas 
               Height          =   3315
               Left            =   75
               TabIndex        =   51
               Top             =   180
               Width           =   9405
               _ExtentX        =   16589
               _ExtentY        =   5847
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   $"Form_Matricula.frx":078E
            End
         End
         Begin VB.ComboBox Cb_EnsinoDisciplinas 
            Height          =   315
            Left            =   840
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   49
            Top             =   540
            Width           =   3555
         End
         Begin VB.Label Label23 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "D I S C I P L I N A S"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   60
            TabIndex        =   53
            Top             =   135
            Width           =   9705
         End
         Begin VB.Label Label31 
            Alignment       =   1  'Right Justify
            Caption         =   "Curso:"
            Height          =   255
            Left            =   240
            TabIndex        =   52
            Top             =   600
            Width           =   555
         End
      End
      Begin VB.Frame Frame_Matricula 
         Height          =   4695
         Left            =   -74910
         TabIndex        =   13
         Top             =   45
         Width           =   9825
         Begin VB.Frame Frame_Matr01 
            Height          =   4155
            Left            =   120
            TabIndex        =   14
            Top             =   480
            Width           =   9600
            Begin VB.Frame Frame8 
               Caption         =   "DISCIPLINA concluida em:"
               Height          =   2325
               Left            =   5400
               TabIndex        =   25
               Top             =   1710
               Width           =   4155
               Begin VB.ComboBox Cb_UFConclusao 
                  Enabled         =   0   'False
                  Height          =   315
                  ItemData        =   "Form_Matricula.frx":0847
                  Left            =   720
                  List            =   "Form_Matricula.frx":089C
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   29
                  Top             =   1395
                  Width           =   750
               End
               Begin VB.TextBox Txt_CidadeConclusao 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   720
                  MaxLength       =   30
                  TabIndex        =   28
                  Top             =   1035
                  Width           =   3255
               End
               Begin VB.CommandButton Bt_GrvConclusao 
                  Caption         =   "Concluir Disciplina"
                  Enabled         =   0   'False
                  Height          =   435
                  Left            =   270
                  TabIndex        =   27
                  Top             =   1800
                  Width           =   3675
               End
               Begin VB.ComboBox Cb_LocConclusao 
                  Enabled         =   0   'False
                  Height          =   315
                  Left            =   720
                  Style           =   2  'Dropdown List
                  TabIndex        =   26
                  Top             =   630
                  Width           =   3270
               End
               Begin MSMask.MaskEdBox Meb_DtConclusao 
                  Height          =   315
                  Left            =   705
                  TabIndex        =   30
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
                  TabIndex        =   34
                  Top             =   1485
                  Width           =   330
               End
               Begin VB.Label Label42 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Cidade:"
                  Height          =   240
                  Left            =   90
                  TabIndex        =   33
                  Top             =   1080
                  Width           =   555
               End
               Begin VB.Label Label28 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Local:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   32
                  Top             =   735
                  Width           =   435
               End
               Begin VB.Label Label25 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Data:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   31
                  Top             =   360
                  Width           =   375
               End
            End
            Begin VB.Frame Frame7 
               Caption         =   "Série:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2835
               Left            =   2760
               TabIndex        =   22
               Top             =   1260
               Width           =   2595
               Begin VB.ListBox Lst_Serie 
                  Enabled         =   0   'False
                  Height          =   2535
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   23
                  Top             =   225
                  Width           =   2400
               End
               Begin VB.Label Lb_SelTodas 
                  Caption         =   "Selecionar Todas"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   1260
                  TabIndex        =   24
                  Top             =   0
                  Width           =   1275
               End
            End
            Begin VB.Frame Frame12 
               Caption         =   "Disciplina:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2775
               Left            =   120
               TabIndex        =   20
               Top             =   1260
               Width           =   2595
               Begin VB.ListBox Lst_Disciplina 
                  Height          =   2400
                  Left            =   90
                  Sorted          =   -1  'True
                  TabIndex        =   21
                  TabStop         =   0   'False
                  Top             =   225
                  Width           =   2400
               End
            End
            Begin VB.Frame Frame13 
               Caption         =   "Curso:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1035
               Left            =   90
               TabIndex        =   17
               Top             =   180
               Width           =   5265
               Begin VB.ComboBox Cb_Ensino 
                  Height          =   315
                  Left            =   1560
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   18
                  Top             =   600
                  Width           =   3525
               End
               Begin MSComCtl2.DTPicker dtpInicio 
                  Height          =   315
                  Left            =   1560
                  TabIndex        =   63
                  Top             =   240
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   60817409
                  CurrentDate     =   39846
               End
               Begin VB.Label Label3 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Data da Matrícula:"
                  Height          =   195
                  Left            =   120
                  TabIndex        =   64
                  Top             =   300
                  Width           =   1395
               End
               Begin VB.Label Label22 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Curso:"
                  Height          =   195
                  Left            =   960
                  TabIndex        =   19
                  Top             =   660
                  Width           =   555
               End
            End
            Begin VB.CommandButton Bt_AlterarMatr 
               Caption         =   "Alterar Matricula em"
               Height          =   735
               Left            =   5715
               Picture         =   "Form_Matricula.frx":090B
               Style           =   1  'Graphical
               TabIndex        =   16
               Top             =   195
               Width           =   3780
            End
            Begin VB.CommandButton Bt_CancelarMatr 
               Caption         =   "Cancelar"
               Enabled         =   0   'False
               Height          =   720
               Left            =   5715
               Picture         =   "Form_Matricula.frx":0C15
               Style           =   1  'Graphical
               TabIndex        =   15
               Top             =   975
               Width           =   3780
            End
         End
         Begin VB.Frame Frame_Matr00 
            Height          =   4095
            Left            =   120
            TabIndex        =   35
            Top             =   480
            Width           =   9600
            Begin VB.Frame Frame19 
               Caption         =   "Série:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2700
               Left            =   2730
               TabIndex        =   44
               Top             =   1350
               Width           =   2595
               Begin VB.ListBox Lst_Serie00 
                  Enabled         =   0   'False
                  Height          =   2310
                  Left            =   90
                  Style           =   1  'Checkbox
                  TabIndex        =   45
                  Top             =   225
                  Width           =   2400
               End
               Begin VB.Label Lb_SelTodasSerie00 
                  Caption         =   "Selecionar Todas"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   1260
                  TabIndex        =   46
                  Top             =   0
                  Width           =   1275
               End
            End
            Begin VB.Frame Frame20 
               Caption         =   "Disciplina:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   2640
               Left            =   90
               TabIndex        =   41
               Top             =   1350
               Width           =   2595
               Begin VB.ListBox Lst_Disciplina00 
                  Enabled         =   0   'False
                  Height          =   2310
                  Left            =   90
                  Sorted          =   -1  'True
                  Style           =   1  'Checkbox
                  TabIndex        =   42
                  TabStop         =   0   'False
                  Top             =   225
                  Width           =   2400
               End
               Begin VB.Label Lb_SelTodasDisc00 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Selecionar Todas"
                  Enabled         =   0   'False
                  Height          =   195
                  Left            =   1260
                  TabIndex        =   43
                  Top             =   0
                  Width           =   1275
               End
            End
            Begin VB.Frame Frame21 
               Caption         =   "Curso:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1155
               Left            =   90
               TabIndex        =   38
               Top             =   180
               Width           =   5265
               Begin VB.ComboBox Cb_Ensino00 
                  Height          =   315
                  Left            =   1500
                  Sorted          =   -1  'True
                  Style           =   2  'Dropdown List
                  TabIndex        =   39
                  Top             =   660
                  Width           =   3705
               End
               Begin MSComCtl2.DTPicker dtpInicio00 
                  Height          =   315
                  Left            =   1500
                  TabIndex        =   65
                  Top             =   240
                  Width           =   1515
                  _ExtentX        =   2672
                  _ExtentY        =   556
                  _Version        =   393216
                  Format          =   60817409
                  CurrentDate     =   39846
               End
               Begin VB.Label Label5 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Data da Matrícula:"
                  Height          =   195
                  Left            =   60
                  TabIndex        =   66
                  Top             =   300
                  Width           =   1395
               End
               Begin VB.Label Label48 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Curso:"
                  Height          =   195
                  Left            =   900
                  TabIndex        =   40
                  Top             =   735
                  Width           =   555
               End
            End
            Begin VB.CommandButton Bt_IncluirMatr00 
               Caption         =   "Incluir Ensino"
               Height          =   720
               Left            =   5715
               Picture         =   "Form_Matricula.frx":0F1F
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   255
               Width           =   3480
            End
            Begin VB.CommandButton Bt_CancelarMatr00 
               Caption         =   "Cancelar"
               Enabled         =   0   'False
               Height          =   720
               Left            =   5715
               Picture         =   "Form_Matricula.frx":1229
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   1020
               Width           =   3480
            End
         End
         Begin VB.Label Label21 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "C A D A S T R O"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   0
            TabIndex        =   47
            Top             =   120
            Width           =   9795
         End
      End
      Begin VB.Frame Frame_Geral 
         Height          =   4695
         Left            =   -74910
         TabIndex        =   9
         Top             =   45
         Width           =   9825
         Begin TabDlg.SSTab sstCurso 
            Height          =   2415
            Left            =   120
            TabIndex        =   68
            Top             =   2160
            Width           =   9615
            _ExtentX        =   16960
            _ExtentY        =   4260
            _Version        =   393216
            TabsPerRow      =   4
            TabHeight       =   520
            TabCaption(0)   =   "Informações"
            TabPicture(0)   =   "Form_Matricula.frx":1533
            Tab(0).ControlEnabled=   -1  'True
            Tab(0).Control(0)=   "Frame1"
            Tab(0).Control(0).Enabled=   0   'False
            Tab(0).ControlCount=   1
            TabCaption(1)   =   "Trancar Curso"
            TabPicture(1)   =   "Form_Matricula.frx":154F
            Tab(1).ControlEnabled=   0   'False
            Tab(1).Control(0)=   "Frame17"
            Tab(1).ControlCount=   1
            TabCaption(2)   =   "Ativação / Renovação"
            TabPicture(2)   =   "Form_Matricula.frx":156B
            Tab(2).ControlEnabled=   0   'False
            Tab(2).Control(0)=   "Frame2"
            Tab(2).ControlCount=   1
            Begin VB.Frame Frame2 
               Height          =   1875
               Left            =   -74820
               TabIndex        =   81
               Top             =   360
               Width           =   9255
               Begin VB.Frame Frame3 
                  Caption         =   "Renovação:"
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
                  Left            =   4620
                  TabIndex        =   87
                  Top             =   120
                  Width           =   4230
                  Begin VB.CommandButton Bt_CancRenovacaoGeral 
                     Caption         =   "Cancelar"
                     Height          =   375
                     Left            =   1695
                     TabIndex        =   89
                     Top             =   720
                     Width           =   1365
                  End
                  Begin VB.CommandButton Bt_AltDtRenovacaoGeral 
                     Caption         =   "Alterar Data"
                     Height          =   375
                     Left            =   255
                     TabIndex        =   88
                     Top             =   720
                     Width           =   1365
                  End
                  Begin MSMask.MaskEdBox Meb_DtRenovacao 
                     Height          =   315
                     Left            =   1560
                     TabIndex        =   90
                     Top             =   240
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
                  Begin VB.Label Label6 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Dt. Renovação:"
                     Height          =   195
                     Left            =   165
                     TabIndex        =   91
                     Top             =   285
                     Width           =   1230
                  End
               End
               Begin VB.Frame Frame18 
                  Caption         =   "Ativação:"
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
                  Left            =   420
                  TabIndex        =   82
                  Top             =   120
                  Width           =   3990
                  Begin VB.CommandButton Bt_AltDtRetornoGeral 
                     Caption         =   "Alterar Data"
                     Height          =   375
                     Left            =   255
                     TabIndex        =   84
                     Top             =   720
                     Width           =   1365
                  End
                  Begin VB.CommandButton Bt_CancRetornoGeral 
                     Caption         =   "Cancelar"
                     Height          =   375
                     Left            =   1695
                     TabIndex        =   83
                     Top             =   720
                     Width           =   1365
                  End
                  Begin MSMask.MaskEdBox Meb_DtRetorno 
                     Height          =   315
                     Left            =   1350
                     TabIndex        =   85
                     Top             =   240
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
                  Begin VB.Label Label49 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Dt. Ativação:"
                     Height          =   195
                     Left            =   165
                     TabIndex        =   86
                     Top             =   285
                     Width           =   1110
                  End
               End
            End
            Begin VB.Frame Frame1 
               Height          =   1875
               Left            =   180
               TabIndex        =   75
               Top             =   360
               Width           =   9255
               Begin VB.Frame Frame16 
                  BeginProperty Font 
                     Name            =   "MS Sans Serif"
                     Size            =   8.25
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   1575
                  Left            =   120
                  TabIndex        =   76
                  Top             =   180
                  Width           =   8955
                  Begin VB.TextBox txtNumCenso 
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   4980
                     MaxLength       =   20
                     TabIndex        =   97
                     Top             =   600
                     Width           =   1575
                  End
                  Begin VB.TextBox txtNumConexao 
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   4980
                     MaxLength       =   20
                     TabIndex        =   96
                     Top             =   240
                     Width           =   1575
                  End
                  Begin VB.TextBox Txt_NumAntGeral 
                     Enabled         =   0   'False
                     Height          =   285
                     Left            =   1740
                     MaxLength       =   15
                     TabIndex        =   92
                     Top             =   660
                     Width           =   1635
                  End
                  Begin VB.CommandButton Bt_AltValCartGeral 
                     Caption         =   "Alterar"
                     Height          =   375
                     Left            =   7260
                     TabIndex        =   78
                     Top             =   300
                     Width           =   1365
                  End
                  Begin VB.CommandButton Bt_CancValCartGeral 
                     Caption         =   "Cancelar"
                     Height          =   375
                     Left            =   7260
                     TabIndex        =   77
                     Top             =   780
                     Width           =   1365
                  End
                  Begin MSComCtl2.DTPicker DTP_ValidadeCard 
                     Height          =   330
                     Left            =   1740
                     TabIndex        =   79
                     Top             =   225
                     Width           =   1590
                     _ExtentX        =   2805
                     _ExtentY        =   582
                     _Version        =   393216
                     Enabled         =   0   'False
                     Format          =   60817409
                     CurrentDate     =   38291
                  End
                  Begin VB.Label Label8 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Num. Conexão:"
                     Height          =   195
                     Left            =   3780
                     TabIndex        =   95
                     Top             =   240
                     Width           =   1155
                  End
                  Begin VB.Label Label7 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Num. Censo:"
                     Height          =   255
                     Left            =   3780
                     TabIndex        =   94
                     Top             =   600
                     Width           =   1155
                  End
                  Begin VB.Label Label39 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Num. matricula antigo:"
                     Height          =   240
                     Left            =   60
                     TabIndex        =   93
                     Top             =   705
                     Width           =   1620
                  End
                  Begin VB.Label Label33 
                     Alignment       =   1  'Right Justify
                     Caption         =   "Carteira valida até:"
                     Height          =   240
                     Left            =   240
                     TabIndex        =   80
                     Top             =   315
                     Width           =   1440
                  End
               End
            End
            Begin VB.Frame Frame17 
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1875
               Left            =   -74820
               TabIndex        =   69
               Top             =   360
               Width           =   9255
               Begin VB.CommandButton Bt_TrancarCursoGeral 
                  Caption         =   "Trancar/ Ativar Curso"
                  Enabled         =   0   'False
                  Height          =   495
                  Left            =   6660
                  TabIndex        =   71
                  Top             =   1140
                  Width           =   2460
               End
               Begin VB.TextBox Txt_MotivoGeral 
                  Enabled         =   0   'False
                  Height          =   285
                  Left            =   675
                  MaxLength       =   100
                  TabIndex        =   70
                  Top             =   630
                  Width           =   8460
               End
               Begin MSComCtl2.DTPicker DTP_TrancaGeral 
                  Height          =   285
                  Left            =   675
                  TabIndex        =   72
                  Top             =   270
                  Width           =   1455
                  _ExtentX        =   2566
                  _ExtentY        =   503
                  _Version        =   393216
                  Enabled         =   0   'False
                  Format          =   60817409
                  CurrentDate     =   39586
               End
               Begin VB.Label Label45 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Data:"
                  Height          =   195
                  Left            =   90
                  TabIndex        =   74
                  Top             =   315
                  Width           =   510
               End
               Begin VB.Label Label46 
                  Caption         =   "Motivo:"
                  Height          =   240
                  Left            =   90
                  TabIndex        =   73
                  Top             =   675
                  Width           =   510
               End
            End
         End
         Begin VB.Frame Frame5 
            Caption         =   "Curso(s):"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1695
            Left            =   60
            TabIndex        =   10
            Top             =   420
            Width           =   9675
            Begin MSFlexGridLib.MSFlexGrid MSFG_EnsinoGeral 
               Height          =   1425
               Left            =   60
               TabIndex        =   11
               Top             =   210
               Width           =   9555
               _ExtentX        =   16854
               _ExtentY        =   2514
               _Version        =   393216
               Cols            =   5
               FixedCols       =   0
               SelectionMode   =   1
               AllowUserResizing=   1
               FormatString    =   $"Form_Matricula.frx":1587
            End
         End
         Begin VB.Label Label40 
            Alignment       =   2  'Center
            BackColor       =   &H00000080&
            Caption         =   "G E R A L"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FFFFFF&
            Height          =   255
            Left            =   45
            TabIndex        =   12
            Top             =   135
            Width           =   9705
         End
      End
   End
   Begin VB.Label lbUnidade 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   780
      TabIndex        =   62
      Top             =   840
      Width           =   6255
   End
   Begin VB.Label Label53 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CADASTRO DE MATRICULA"
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
      TabIndex        =   7
      Top             =   0
      Width           =   10515
   End
   Begin VB.Label Lb_AtivoInativo 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   8730
      TabIndex        =   6
      Top             =   795
      Width           =   1635
   End
   Begin VB.Label Label47 
      Caption         =   "Status da Matricula:"
      Height          =   195
      Left            =   7200
      TabIndex        =   5
      Top             =   885
      Width           =   1410
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Unidade:"
      Height          =   195
      Left            =   60
      TabIndex        =   3
      Top             =   855
      Width           =   675
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Matricula:"
      Height          =   195
      Left            =   7920
      TabIndex        =   2
      Top             =   495
      Width           =   690
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Nome:"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   450
      Width           =   495
   End
End
Attribute VB_Name = "Form_Matricula"
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






Private Function Chk_LstDisciplina00() As Boolean
    Dim cDisciplina As Integer
    'Checa se existe alguma opcao na lista de Disciplina00 foi selecionada
    For cDisciplina = 0 To Lst_Disciplina00.ListCount - 1
        If Lst_Disciplina00.Selected(cDisciplina) = True Then
            Chk_LstDisciplina00 = True
            Exit Function
        End If
        Chk_LstDisciplina00 = False
    Next
End Function

Private Function Chk_LstSerie00() As Boolean
    Dim cSerie As Integer
    'Checa se existe alguma opcao na lista de serie00 foi selecionada
    For cSerie = 0 To Lst_Serie00.ListCount - 1
        If Lst_Serie00.Selected(cSerie) = True Then
            Chk_LstSerie00 = True
            Exit Function
        End If
        Chk_LstSerie00 = False
    Next
End Function




Private Function pgObsNH(ID As String)
    On Error GoTo TreErroPgObs
    Dim Rs As Recordset
    If Trim(ID) = "" Then
        pgObsNH = " "
        Exit Function
    End If
    Set Rs = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MPID = " & ID & " ORDER BY Seq")
    If Rs.BOF And Rs.EOF Then
            pgObsNH = " "
        Else
            Rs.MoveLast
            pgObsNH = IIf(IsNull(Rs.Fields("Obs")), " ", Rs.Fields("Obs"))
    End If
    Rs.Close
    Exit Function
TreErroPgObs:
    pgObsNH = ""
    Resume Next
End Function

Private Sub Bt_AltDtRetornoGeral_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    
    If Not IsNumeric(MatrID) Then Exit Sub
    
    Select Case left(Bt_AltDtRetornoGeral.Caption, 1)
        Case "A"
            Meb_DtRetorno.Enabled = True
            Bt_AltDtRetornoGeral.Caption = "Gravar Data"
        Case "G"
            Meb_DtRetorno.Enabled = False
            Bt_AltDtRetornoGeral.Caption = "Alterar Data"
            
            Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrId = '" & MatrID & "'")
            If RsMatricula.BOF And RsMatricula.EOF Then
                    MsgBox "Erro ao localizar a Matricula n." & MatrID & vbCrLf & "Operação cancelada.", vbInformation, "CESNet - Aviso!"
                    Exit Sub
                Else
                    RsMatricula.MoveFirst
                    RsMatricula.Edit
                    
                    Meb_DtRetorno.PromptInclude = False
                    If Trim(Meb_DtRetorno.Text) = "" Then
                            RsMatricula.Fields("DtRetorno") = Null
                            Meb_DtRetorno.PromptInclude = True
                        Else
                            Meb_DtRetorno.PromptInclude = True
                            RsMatricula.Fields("DtRetorno") = Trim(Meb_DtRetorno.Text)
                            ExcluirAviso 1
                            GrvRetorno 1, Meb_DtRetorno.Text
                    End If
                    
                    RsMatricula.Update
                    MstDadosAluno
            End If
    End Select
End Sub
Private Sub Bt_AltDtRenovacaoGeral_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    
    If Not IsNumeric(MatrID) Then Exit Sub
    
    Select Case left(Bt_AltDtRenovacaoGeral.Caption, 1)
        Case "A"
            Meb_DtRenovacao.Enabled = True
            Bt_AltDtRenovacaoGeral.Caption = "Gravar Data"
        Case "G"
            Meb_DtRenovacao.Enabled = False
            Bt_AltDtRenovacaoGeral.Caption = "Alterar Data"
            
            'Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrId = '" & MatrID & "'")
            'If RsMatricula.BOF And RsMatricula.EOF Then
            '        MsgBox "Erro ao localizar a Matricula n." & MatrID & vbCrLf & "Operação cancelada.", vbInformation, "CESNet - Aviso!"
            '        Exit Sub
            '    Else
            '        RsMatricula.MoveFirst
            '        RsMatricula.Edit
            '
            '        Meb_DtRenovacao.PromptInclude = False
            '        If Trim(Meb_DtRenovacao.Text) = "" Then
            '                RsMatricula.Fields("DtRenovacao") = Null
            '                Meb_DtRenovacao.PromptInclude = True
            '            Else
            '                Meb_DtRenovacao.PromptInclude = True
            '                RsMatricula.Fields("DtRenovacao") = Trim(Meb_DtRenovacao.Text)
                            ExcluirAviso 2
                            GrvRetorno 2, Meb_DtRenovacao.Text
            '        End If
                    
            '        RsMatricula.Update
            '        MstDadosAluno
            'End If
    End Select
End Sub
Private Sub ExcluirAviso(tpAviso As Integer)
    If Trim(MatrID) = "" Then Exit Sub
    BD.Execute "DELETE * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "' AND Codigo=" & tpAviso
    Call RegLog(MatrID, "Aviso de bloqueio codigo " & tpAviso & " removido automat. pela tela de Matricula")
End Sub
Private Sub GrvRetorno(iTipo As Integer, dDt As Date)
    '#################################################################
    '### Data: 15/09/2011
    '### Grava o Ativacao/Retorno do aluno caso esteja inativo
    '### Onde:
    '### 1 - Ativacao = Coloca o aluno ativo no sistema
    '### 2 - Retorno = Marca a data que o aluno fez o ultimo retorno
    '#################################################################
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM MatriculaRetorno WHERE MatrID = '" & MatrID & "' AND DtRetorno = #" & Format(dDt, "MM/DD/YYYY") & "# AND TpMov = " & iTipo
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            Rst.AddNew
            Rst.Fields("MatrID") = MatrID
            Rst.Fields("DtRetorno") = dDt
            Rst.Fields("TpMov") = iTipo
            Rst.Fields("UsuarioID") = UsuarioID
            Rst.Fields("DtHr") = Now
            Rst.Update
        Else
            MsgBox "Retorno ja cadastrado. Ação CANCELADA!", vbInformation, "Aviso"
            
    End If
    Rst.Close
    
End Sub
Private Sub Bt_AlterarMatr_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub

    If ChkAcesso(Me.Name, "N") = False Then Exit Sub

    
    If Trim(Lst_Disciplina.Text) = "" Then Exit Sub
    Select Case left(Bt_AlterarMatr.Caption, 1)
        Case "A"
            Lst_Disciplina.Enabled = False
            Lst_Serie.Enabled = True
            HDConclDiscipl (False)
            Bt_AlterarMatr.Caption = "Gravar em " & Lst_Disciplina.Text
            Bt_CancelarMatr.Enabled = True
            Cb_Ensino.Enabled = False
            dtpInicio.Enabled = False
            Lb_SelTodas.Enabled = True
        Case "G"
            Lb_SelTodas.Enabled = False
            Call GrvSeries
            Call Bt_CancelarMatr_Click
    End Select
End Sub





Private Sub Bt_AltValCartGeral_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub

    
    If Not IsNumeric(MatrID) Then Exit Sub
    Select Case left(Bt_AltValCartGeral.Caption, 1)
        Case "A"
            DTP_ValidadeCard.Enabled = True
            Txt_NumAntGeral.Enabled = True
            txtNumConexao.Enabled = True
            txtNumCenso.Enabled = True
            Bt_AltValCartGeral.Caption = "Gravar"
        Case "G"
            DTP_ValidadeCard.Enabled = False
            Txt_NumAntGeral.Enabled = False
            txtNumConexao.Enabled = False
            txtNumCenso.Enabled = False
            
            Bt_AltValCartGeral.Caption = "Alterar"
            
            Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrId = '" & MatrID & "'")
            If RsMatricula.BOF And RsMatricula.EOF Then
                    MsgBox "Erro ao localizar a Matricula n." & MatrID & vbCrLf & "Operação cancelada.", vbInformation, "CESNet - Aviso!"
                    Exit Sub
                Else
                    RsMatricula.MoveFirst
                    RsMatricula.Edit
                    RsMatricula.Fields("ValCard") = DTP_ValidadeCard.Value
                    RsMatricula.Fields("NumAnt") = IIf(Trim(Txt_NumAntGeral.Text) = "", Null, Trim(Txt_NumAntGeral.Text))
                    
                    RsMatricula.Fields("NumConexao") = IIf(Trim(txtNumConexao.Text) = "", Null, Trim(txtNumConexao.Text))
                    RsMatricula.Fields("NumCenso") = IIf(Trim(txtNumCenso.Text) = "", Null, Trim(txtNumCenso.Text))
                    RsMatricula.Update
            End If
    End Select

End Sub



Private Sub Bt_CancelarMatr_Click()
    Lst_Disciplina.Enabled = True
    Lst_Serie.Enabled = False
    Lb_SelTodas.Enabled = False
    Bt_AlterarMatr.Caption = "Alterar Matricula em " & Lst_Disciplina.Text
    Bt_CancelarMatr.Enabled = False
    Cb_Ensino.Enabled = True
    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND Disciplinaid = " & DisciplinaID & " AND Aprovado = False")
    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
            HDConclDiscipl (True)
        Else
            HDConclDiscipl (False)
    End If
    RsMatriculaSerie.Close
End Sub

Private Sub Bt_CancelarMatr00_Click()
    Lst_Disciplina00.Enabled = False
    Lst_Serie00.Enabled = False
    Cb_Ensino00.Enabled = True
    
    Lb_SelTodasDisc00.Enabled = False
    Lb_SelTodasSerie00.Enabled = False
            
    Lst_Disciplina00.Clear
    Lst_Serie00.Clear
    Cb_Ensino00.Clear
    
    Bt_IncluirMatr00.Caption = "Incluir Ensino " & Cb_Ensino00.Text
    Bt_IncluirMatr00.Enabled = True
    Bt_CancelarMatr00.Enabled = False
End Sub


Private Sub Bt_CancRetornoGeral_Click()
    Meb_DtRetorno.Enabled = False
    Bt_AltDtRetornoGeral.Caption = "Alterar Data"
End Sub

Private Sub Bt_CancValCartGeral_Click()
    DTP_ValidadeCard.Enabled = False
    Txt_NumAntGeral.Enabled = False
    txtNumConexao.Enabled = False
    txtNumCenso.Enabled = False
    Bt_AltValCartGeral.Caption = "Alterar"
    
    
End Sub












Private Sub Bt_IncluirMatr00_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    
    If Trim(Cb_Ensino00.Text) = "" Then
        MsgBox "Favor selecionar um ensino!", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    
    Select Case left(Bt_IncluirMatr00.Caption, 1)
        Case "I"
            Lb_SelTodasDisc00.Enabled = True
            Lb_SelTodasSerie00.Enabled = True
            Lst_Disciplina00.Enabled = True
            Lst_Serie00.Enabled = True
            Cb_Ensino00.Enabled = False
            
            'Bt_IncluirMatr00.Enabled = False
            Bt_CancelarMatr00.Enabled = True
            
            Bt_IncluirMatr00.Caption = "Gravar Ensino " & Cb_Ensino00.Text
        Case "G"
            If Chk_LstDisciplina00 = False Then
                MsgBox "Favor selecionar um item da listagem de disciplinas.", vbInformation, "CESNet - Aviso"
                Exit Sub
            End If
            If Chk_LstSerie00 = False Then
                MsgBox "Favor selecionar um item da listagem de séries.", vbInformation, "CESNet - Aviso"
                Exit Sub
            End If
            
            'Checa se o ensino esta concluido
            Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
            If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                Else
                    RsMatriculaEnsino.MoveFirst
                    If IsNull(RsMatriculaEnsino.Fields("DtFinal")) = False Then
                        MsgBox "Ensino já concluido. Favor selecione outro.", vbInformation, "CESNet - Aviso"
                        Exit Sub
                    End If
            End If
            
            Cb_Ensino00.Enabled = True
            Lst_Disciplina00.Enabled = False
            Lst_Serie00.Enabled = False
            Lb_SelTodasDisc00.Enabled = False
            Lb_SelTodasSerie00.Enabled = False
            
            Call GrvEnsino00
            
            Bt_IncluirMatr00.Caption = "Incluir Ensino " ' & Cb_Ensino00.Text
            Bt_IncluirMatr00.Enabled = True
            Bt_CancelarMatr00.Enabled = False
            
            Frame_Matr01.Visible = True
            Frame_Matr00.Visible = False
            Cb_Ensino.AddItem (PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)))
            Cb_Ensino.Text = PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False))
    End Select
End Sub




Private Sub Bt_TrancarCursoGeral_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    Dim CursoID     As Integer
    Dim RsMatrEns   As Recordset
    If Trim(Txt_MotivoGeral.Text) = "" Then
        MsgBox "O campo MOTIVO não pode ser um campo nulo.", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    CursoID = PgIDEnsino(MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 0))
    Set RsMatrEns = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & CursoID)
    If RsMatrEns.BOF And RsMatrEns.EOF Then
            MsgBox "Erro ao localizar Curso", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsMatrEns.MoveFirst
            If RsMatrEns.Fields("Trancado") = False Then
                    If MsgBox("Deseja realmente TRANCAR a matricula?", vbYesNo + vbInformation, "CESNet - Aviso") = vbYes Then
                        RsMatrEns.Edit
                        RsMatrEns.Fields("Trancado") = True
                        RsMatrEns.Fields("DtFinal") = Trim(DTP_TrancaGeral.Value)
                        RsMatrEns.Fields("Local") = Trim(Txt_MotivoGeral.Text)
                        RsMatrEns.Fields("UsuarioID") = UsuarioID
                        RsMatrEns.Fields("DtHrSis") = Now()
                        RsMatrEns.Update
                        MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 2) = DTP_TrancaGeral.Value
                        MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 3) = Trim(Txt_MotivoGeral.Text)
                        DTP_TrancaGeral.Enabled = False
                        Txt_MotivoGeral.Text = ""
                        Bt_TrancarCursoGeral.Enabled = False
                        RegLog MatrID, "Trancou o curso " & PgNomeEnsino(CursoID)
                    End If
                Else
                    If MsgBox("Deseja realmente DESTRANCAR a matricula?", vbYesNo + vbInformation, "CESNet - Aviso") = vbYes Then
                        'Checa se nao ha outro curso em abeto
                        If PgMatrEnsino(MatrID) <> 0 Then
                            MsgBox "Curso nao pode ser destrancado devido haver outro em andamento.", vbInformation, "CESNet - Aviso"
                            RsMatrEns.Close
                            Exit Sub
                        End If
                        'Checa se o curso foi concluido ou trancaddo
                        'If RsMatrEns.Fields("Trancado") <> 0 Then
                        '    MsgBox "Este curso não foi trancado e sim concluido pelo sistema. Ação cancelada", vbInformation, "CESNet - Aviso"
                        '    RsMatrEns.Close
                        '    Exit Sub
                        'End If
                        RsMatrEns.Edit
                        RsMatrEns.Fields("Trancado") = False
                        RsMatrEns.Fields("DtFinal") = Null
                        RsMatrEns.Fields("Local") = Null
                        RsMatrEns.Fields("UsuarioID") = UsuarioID
                        RsMatrEns.Fields("DtHrSis") = Now()
                        RsMatrEns.Update
                        MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 2) = ""
                        MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 3) = ""
                        DTP_TrancaGeral.Enabled = False
                        Txt_MotivoGeral.Text = ""
                        Bt_TrancarCursoGeral.Enabled = False
                        RegLog MatrID, "Ativou o curso " & PgNomeEnsino(CursoID)
                    End If
            End If
            RsMatrEns.Close
    
    End If
End Sub









Private Sub btExcluirDisciplina_Click()
    If ChkAcesso(Me.Name, "E") = False Then Exit Sub
 
    Dim DisciplinaID As Integer
    Dim RsTMP As Recordset
    
    
    DisciplinaID = PgIDDisciplina(MSFG_HstDisciplinas.TextMatrix(MSFG_HstDisciplinas.Row, 1))
    If DisciplinaID = 0 Then Exit Sub
    If MsgBox("Deseja realmente EXCLUIR a disciplina de " & PgNomeDisciplina(DisciplinaID) & "?", vbYesNo + vbInformation, "CESNet - Aviso") = vbYes Then
        'Autentica o usuario
        If Form_AutenticacaoUsuario.CarregarForm = True Then
            Call RegLog(MatrID, "Excluiu a disciplina " & PgNomeEnsino(EnsinoID) & "/" & PgNomeDisciplina(DisciplinaID))
            BD.Execute "DELETE * FROM MatriculaDisciplina WHERE MatrID ='" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID
            BD.Execute "DELETE * FROM MatriculaSerie WHERE MatrID ='" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID
            
            
            'Lista as provas Exclidas
            Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID ='" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
            If RsTMP.BOF And RsTMP.EOF Then
                    Call RegLog(MatrID, "Tela: MATRICULA/Disciplinas - Exclusao da disciplina (sem provas aplicadas): " & PgNomeEnsino(EnsinoID) & " / " & PgNomeDisciplina(DisciplinaID))
                Else
                    RsTMP.MoveFirst
                    Do Until RsTMP.EOF
                        Call RegLog(MatrID, "Tela: MATRICULA/Disciplinas - Exclusao da disciplina: " & PgNomeEnsino(EnsinoID) & "/" & PgNomeDisciplina(DisciplinaID) & _
                                    " / " & RsTMP.Fields("NProva") & " / " & RsTMP.Fields("Tipo") & " / " & RsTMP.Fields("DtAvaliacao") & _
                                    " / " & RsTMP.Fields("Status") & " / " & RsTMP.Fields("ProfIDN"))
                        RsTMP.MoveNext
                    Loop
            End If
            RsTMP.Close
            BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID ='" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID
            Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID=" & EnsinoID)
            If RsTMP.BOF And RsTMP.EOF Then
                BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID
                Call RegLog(MatrID, "Curso " & PgNomeEnsino(EnsinoID) & " EXCLUIDO devido nao haver disciplinas")
            End If
            
            Cb_EnsinoDisciplinas_Click
            
        End If
    End If
   
End Sub

Private Sub Cb_Ensino00_Click()
    If Trim(Cb_Ensino00.Text) = "" Then
        Lst_Disciplina00.Clear
        Lst_Serie00.Clear
        Bt_IncluirMatr00.Caption = "Incluir Ensino"
    End If
    Bt_IncluirMatr00.Caption = "Incluir Ensino " & Cb_Ensino00.Text
    LstDisciplinas00
End Sub

Private Sub Cb_Ensino00_DropDown()
    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Then
            MebMatricula.PromptInclude = True
            Exit Sub
        Else
            MebMatricula.PromptInclude = True
    End If
    Cb_Ensino00.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.BOF Then
            MsgBox "Não existe nenhum Ensino cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_Ensino00.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If

End Sub

Private Sub Cb_EnsinoDisciplinas_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub
Private Sub Cb_EnsinoProvas_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Cb_DisciplinaProvas_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub




Private Sub Cb_LocConclusao_Click()
    If Trim(Cb_LocConclusao.Text) = "" Then
            Txt_CidadeConclusao.Text = ""
            Cb_UFConclusao.Text = " "
            Exit Sub
        Else
            If IsNumeric(Trim(left(Cb_LocConclusao.Text, 3))) Then
                Txt_CidadeConclusao.Text = PgDadosInstEns(Trim(left(Cb_LocConclusao.Text, 3))).Cidade
                Cb_UFConclusao.Text = PgDadosInstEns(Trim(left(Cb_LocConclusao.Text, 3))).UF
            End If
    End If
End Sub

Private Sub Cb_LocConclusao_DropDown()
    Cb_LocConclusao.Clear
    Set RsInstEnsino = BD.OpenRecordset("SELECT * FROM InstEnsino ORDER BY Descr")
    If RsInstEnsino.BOF And RsInstEnsino.EOF Then
            Cb_LocConclusao.Clear
        Else
            RsInstEnsino.MoveFirst
            Do Until RsInstEnsino.EOF
                Cb_LocConclusao.AddItem (left(String(3, "0"), 3 - Len(Trim(RsInstEnsino.Fields("ID"))))) & Trim(RsInstEnsino.Fields("ID")) & " - " & _
                                         Trim(RsInstEnsino.Fields("Descr"))
                RsInstEnsino.MoveNext
            Loop
    End If
End Sub

Private Sub Cb_LocConclusao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
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


Private Sub Cb_UFConclusao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub







Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub







Private Sub Bt_GrvConclusao_Click()
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    
    Dim tmp As String
    
    
    Meb_DtConclusao.Text = Trim(Meb_DtConclusao.Text)
    'Cb_LocConclusao.Text = Trim(Cb_LocConclusao.Text)
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
    'Checa se existe Ensino iniciado ************ ERRO DUPLICANDO ENSINO  *********************
    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND IsNull(DtFinal)")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.AddNew
            RsTMP.Fields("MatrID") = MatrID
            RsTMP.Fields("EnsinoID") = EnsinoID
            RsTMP.Update
        Else
            'Checa se existe ensino aberto que nao seja o editado pelo usuario
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                If EnsinoID = RsTMP.Fields("EnsinoID") Then
                    Else
                        MsgBox "Esta matricula já possui o curso " & PgNomeEnsino(RsTMP.Fields("EnsinoID")) & " em andamento.", vbInformation, "CESNet - Atenção"
                        'Lst_Serie.Selected(Lst_Serie.ListIndex) = False
                        Call LimpConclDiscipl
                        'Meb_DtConclusao.PromptInclude = False
                        'Meb_DtConclusao.Text = ""
                        'Meb_DtConclusao.PromptInclude = True
                        'Cb_LocConclusao.Clear
                        Exit Sub
                End If
                RsTMP.MoveNext
            Loop
    End If
    DisciplinaID = PgIDDisciplina(Lst_Disciplina.Text)
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            RsMatriculaDisciplina.AddNew
            RsMatriculaDisciplina.Fields("MatrID") = MatrID
            RsMatriculaDisciplina.Fields("EnsinoID") = EnsinoID
            RsMatriculaDisciplina.Fields("DisciplinaID") = DisciplinaID
            Meb_DtConclusao.PromptInclude = False
            If Trim(Meb_DtConclusao.Text) = "" Then
                Else
                    Meb_DtConclusao.PromptInclude = True
                    RsMatriculaDisciplina.Fields("DtInicio") = Meb_DtConclusao.Text
                    RsMatriculaDisciplina.Fields("DtConclusao") = Meb_DtConclusao.Text
            End If
            Meb_DtConclusao.PromptInclude = True
            'RsMatriculaDisciplina.Fields("Local") = IIf(Cb_LocConclusao.Text = "", Null, Cb_LocConclusao.Text)
            'RsMatriculaDisciplina.Fields("Cidade") = Trim(Txt_CidadeConclusao.Text)
            'RsMatriculaDisciplina.Fields("UF") = Trim(Cb_UFConclusao.Text)
            RsMatriculaDisciplina.Fields("InstID") = left(Trim(Cb_LocConclusao.Text), 3)
            tmp = PgDadosInstEns(left(Trim(Cb_LocConclusao.Text), 3)).Nome
            RsMatriculaDisciplina.Fields("Local") = IIf(Trim(tmp) = "", Null, tmp)
            tmp = PgDadosInstEns(left(Trim(Cb_LocConclusao.Text), 3)).Abreviatura
            RsMatriculaDisciplina.Fields("Abrev") = IIf(Trim(tmp) = "", Null, tmp)
            RsMatriculaDisciplina.Fields("Cidade") = IIf(Trim(Txt_CidadeConclusao.Text) = "", Null, Trim(Txt_CidadeConclusao.Text))
            RsMatriculaDisciplina.Fields("UF") = IIf(Trim(Cb_UFConclusao.Text) = "", Null, Trim(Cb_UFConclusao.Text))
            RsMatriculaDisciplina.Fields("UsuarioID") = UsuarioID
            RsMatriculaDisciplina.Fields("DtHrSis") = Now()
            RsMatriculaDisciplina.Update
            MsgBox "Disciplina concluida com sucesso.", vbInformation, "CESNet - Atenção"
        Else
            If Meb_DtConclusao.Text = "" And Cb_LocConclusao.Text = "" Then
                    RsMatricula.Delete
                Else
                    
                    
                    RsMatriculaDisciplina.Edit
                    RsMatriculaDisciplina.Fields("DtInicio") = IIf(Meb_DtConclusao.Text = "", Null, Meb_DtConclusao.Text)
                    RsMatriculaDisciplina.Fields("DtConclusao") = IIf(Meb_DtConclusao.Text = "", Null, Meb_DtConclusao.Text)
                    If IsNumeric(left(Trim(Cb_LocConclusao.Text), 3)) Then
                            RsMatriculaDisciplina.Fields("InstID") = left(Trim(Cb_LocConclusao.Text), 3)
                            tmp = PgDadosInstEns(left(Trim(Cb_LocConclusao.Text), 3)).Nome
                            RsMatriculaDisciplina.Fields("Local") = IIf(Trim(tmp) = "", Null, tmp)
                            tmp = PgDadosInstEns(left(Trim(Cb_LocConclusao.Text), 3)).Abreviatura
                            RsMatriculaDisciplina.Fields("Abrev") = IIf(Trim(tmp) = "", Null, tmp)
                        Else
                            RsMatriculaDisciplina.Fields("InstID") = 0
                            tmp = Trim(Cb_LocConclusao.Text)
                            RsMatriculaDisciplina.Fields("Local") = IIf(Trim(tmp) = "", Null, tmp)
                            tmp = left(Trim(Cb_LocConclusao.Text), 10)
                            RsMatriculaDisciplina.Fields("Abrev") = IIf(Trim(tmp) = "", Null, tmp)
                    End If
                    RsMatriculaDisciplina.Fields("Cidade") = IIf(Trim(Txt_CidadeConclusao.Text) = "", Null, Trim(Txt_CidadeConclusao.Text))
                    RsMatriculaDisciplina.Fields("UF") = IIf(Trim(Cb_UFConclusao.Text) = "", Null, Trim(Cb_UFConclusao.Text))
                    RsMatriculaDisciplina.Fields("UsuarioID") = UsuarioID
                    RsMatriculaDisciplina.Fields("DtHrSis") = Now()
                    RsMatriculaDisciplina.Update
                    MsgBox "Disciplina alterada e concluida com sucesso.", vbInformation, "CESNet - Aviso"
            End If
    End If
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
                RsMatriculaEnsino.Fields("UsuarioID") = UsuarioID
                RsMatriculaEnsino.Fields("DtHrSis") = Now()
                RsMatriculaEnsino.Update
            Else
                RsMatriculaDisciplina.MoveLast
                'RsMatriculaEnsino.MoveFirst
                RsMatriculaEnsino.Edit
                RsMatriculaEnsino.Fields("DtFinal") = RsMatriculaDisciplina.Fields("DtConclusao")
                RsMatriculaEnsino.Fields("Local") = RsMatriculaDisciplina.Fields("Local")
                RsMatriculaEnsino.Fields("UsuarioID") = UsuarioID
                RsMatriculaEnsino.Fields("DtHrSis") = Now()
                RsMatriculaEnsino.Update
        End If
    End If
End Sub







Private Sub Cb_Ensino_Click()
    Lst_Serie.Clear
    LstDisciplinas
End Sub

Private Sub Cb_Ensino_DropDown()
    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Then
            MebMatricula.PromptInclude = True
            Exit Sub
        Else
            MebMatricula.PromptInclude = True
    End If
    Cb_Ensino.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.BOF Then
            MsgBox "Não existe nenhum Ensino cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_Ensino.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If
End Sub
Private Sub Cb_Ensino_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
            KeyAscii = 0
    End If
End Sub


Private Sub Cb_EnsinoDisciplinas_Click()
    btExcluirDisciplina.Enabled = False
    If Trim(Cb_EnsinoDisciplinas.Text) = "" Then
        Exit Sub
    End If
    'LISTA TODAS AS DisciplinaS DO ALUNO de acordo com Ensino
    'Set RsEnsino = BD.OpenRecordset("Select * FROM Ensino WHERE Descr = '" & Cb_EnsinoDisciplinas.Text & "'")
    EnsinoID = PgIDEnsino(Cb_EnsinoDisciplinas.Text) 'RsEnsino.Fields("EnsinoID")
    Set RsGradeEnsinoDisciplinas = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    With RsGradeEnsinoDisciplinas
        If .BOF And .EOF Then
            MSFG_HstDisciplinas.Rows = 1
            MSFG_HstDisciplinas.Rows = 2
            Exit Sub
        End If
        lin = 1
        MSFG_HstDisciplinas.Rows = 1
        MSFG_HstDisciplinas.Rows = 2
        .MoveFirst
        Do Until .EOF
            'Set RsEnsino = BD.OpenRecordset("Select * FROM Ensino WHERE EnsinoID = " & .Fields("EnsinoID"))
            Ensino = PgNomeEnsino(.Fields("EnsinoID")) 'RsEnsino.Fields("Descr")
            'Set RsDisciplina = BD.OpenRecordset("Select * FROM Disciplina WHERE DisciplinaID = " & .Fields("DisciplinaID"))
            Disciplina = PgNomeDisciplina(.Fields("DisciplinaID")) 'RsDisciplina.Fields("Descr")
            DoEvents
            MSFG_HstDisciplinas.TextMatrix(lin, 0) = Ensino
            MSFG_HstDisciplinas.TextMatrix(lin, 1) = Disciplina
            MSFG_HstDisciplinas.TextMatrix(lin, 2) = IIf(IsNull(.Fields("DtInicio")), " ", .Fields("DtInicio"))
            MSFG_HstDisciplinas.TextMatrix(lin, 3) = IIf(IsNull(.Fields("DtConclusao")), " ", .Fields("DtConclusao"))
            MSFG_HstDisciplinas.TextMatrix(lin, 4) = IIf(IsNull(.Fields("Local")), " ", .Fields("local"))
            .MoveNext
            lin = lin + 1
            MSFG_HstDisciplinas.Rows = MSFG_HstDisciplinas.Rows + 1
        Loop
        MSFG_HstDisciplinas.Rows = MSFG_HstDisciplinas.Rows - 1
    End With
End Sub

Private Sub Cb_EnsinoDisciplinas_DropDown()
    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Then
            MebMatricula.PromptInclude = True
            Exit Sub
        Else
            MebMatricula.PromptInclude = True
    End If
    Cb_EnsinoDisciplinas.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.BOF Then
            MsgBox "Não existe nenhum Ensino cadastrado. Por favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_EnsinoDisciplinas.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If
End Sub


Private Sub Cb_EnsinoProvas_Click()
    EnsinoID = PgIDEnsino(Cb_EnsinoProvas.Text)
    'Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino WHERE Descr = '" & Cb_EnsinoProvas.Text & "'")
    'If RsEnsino.BOF And RsEnsino.EOF Then
    '    Else
    '        RsEnsino.MoveFirst
    '        EnsinoID = RsEnsino.Fields("EnsinoID")
    'End If
End Sub

Private Sub Cb_EnsinoProvas_DropDown()
    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Then
            MebMatricula.PromptInclude = True
            Exit Sub
        Else
            MebMatricula.PromptInclude = True
    End If
    
    Cb_EnsinoProvas.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.BOF Then
            MsgBox "Não existe nenhum Ensino cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_EnsinoProvas.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If
End Sub

Private Sub Cb_EnsinoProvas_GotFocus()
    Cb_DisciplinaProvas.Clear
    MSFG_HstProvas.Rows = 1
    MSFG_HstProvas.Rows = 2
    
End Sub

Private Sub Cb_DisciplinaProvas_Click()
    'On Error GoTo TrtErroPrv
    DisciplinaID = PgIDDisciplina(Cb_DisciplinaProvas.Text) 'RsDisciplina.Fields("DisciplinaID")
    MSFG_HstProvas.Rows = 1
    MSFG_HstProvas.Rows = 2
    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
        MsgBox "Não existe nenhuma Prova cadastrada para esta Matricula.", vbInformation, "CESNet - Aviso!"
        MSFG_HstProvas.Rows = 1
        MSFG_HstProvas.Rows = 2
        Exit Sub
    End If
    RsMatriculaSerie.MoveFirst
    lin = 1
    Do Until RsMatriculaSerie.EOF
        
        SerieID = RsMatriculaSerie.Fields("SerieID")
      
            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID=" & DisciplinaID & " AND SerieID=" & SerieID & " ORDER BY NProva, ModuloID")

            If RsProvas.BOF And RsProvas.EOF Then
                    MSFG_HstProvas.Rows = 1
                    MSFG_HstProvas.Rows = 2
                Else
                    RsProvas.MoveFirst
                    Do Until RsProvas.EOF
                        MSFG_HstProvas.TextMatrix(lin, 0) = PgNomeModulo(IIf(IsNull(RsProvas.Fields("ModuloID")), 0, RsProvas.Fields("ModuloID")))
                        MSFG_HstProvas.TextMatrix(lin, 1) = IIf(IsNull(RsProvas.Fields("NProva")), " ", RsProvas.Fields("NProva"))
                        'MSFG_HstProvas.TextMatrix(lin, 2) = IIf(IsNull(RsProvas.Fields("Assunto")), " ", RsProvas.Fields("Assunto"))
                        MSFG_HstProvas.TextMatrix(lin, 3) = IIf(IsNull(RsProvas.Fields("Pag")), " ", RsProvas.Fields("Pag"))
                        
                        'Listas as provas sem o uso da RefTrafegoID
                        Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND NProva = '" & RsProvas.Fields("NProva") & "' AND RefTrafegoID = " & RsProvas.Fields("RefTrafegoID"))
                        'Listas as provas usando da RefTrafegoID
                        'Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " AND NProva = '" & RsProvas.Fields("NProva") & "'")
                        If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
                                RsMatriculaProva.Close
                            Else
                                RsMatriculaProva.MoveFirst
                                MSFG_HstProvas.TextMatrix(lin, 2) = IIf(IsNull(RsMatriculaProva.Fields("Assunto")), " ", RsMatriculaProva.Fields("Assunto"))
                                'MSFG_HstProvas.TextMatrix(lin, 3) = IIf(IsNull(RsMatriculaProva.Fields("Pag")), " ", RsMatriculaProva.Fields("Pag"))
                                MSFG_HstProvas.TextMatrix(lin, 4) = IIf(IsNull(RsMatriculaProva.Fields("DtAvaliacao")), " ", RsMatriculaProva.Fields("DtAvaliacao"))
                                MSFG_HstProvas.TextMatrix(lin, 5) = IIf(IsNull(RsMatriculaProva.Fields("Tipo")), " ", RsMatriculaProva.Fields("Tipo"))
                                If SisNota = True Then
                                        MSFG_HstProvas.TextMatrix(lin, 6) = IIf(IsNull(RsMatriculaProva.Fields("Nota")), " ", RsMatriculaProva.Fields("Nota"))

                                    Else
                                        Select Case IIf(IsNull(RsMatriculaProva.Fields("Nota")), 0, RsMatriculaProva.Fields("Nota"))
                                            Case Is >= Val(NotaMedia)
                                                MSFG_HstProvas.TextMatrix(lin, 6) = "SIM"
                                            Case Is < Val(NotaMedia)
                                                MSFG_HstProvas.TextMatrix(lin, 6) = "NÃO"
                                            Case Else
                                                MSFG_HstProvas.TextMatrix(lin, 6) = " "
                                        End Select
                                End If
                                MSFG_HstProvas.TextMatrix(lin, 7) = IIf(IsNull(RsMatriculaProva.Fields("Obs")), " ", RsMatriculaProva.Fields("Obs"))
                                
                                Dim XXX As String
                                XXX = IIf(IsNull(RsMatriculaProva.Fields("Status")), "0", RsMatriculaProva.Fields("Status"))
                                
                                Select Case XXX 'RsMatriculaProva.Fields("Nota")
                                    Case "HB" 'Is >= Val(NotaMedia)
                                        MSFG_HstProvas.Row = MSFG_HstProvas.Rows - 1
                                        'MSFG_HstProvas.RowSel = MSFG_HstProvas.Rows - 1
                                        MSFG_HstProvas.Col = 0
                                        MSFG_HstProvas.ColSel = MSFG_HstProvas.Cols - 1
                                        MSFG_HstProvas.FillStyle = flexFillRepeat
                                        MSFG_HstProvas.CellForeColor = QBColor(9)
                                    Case "NH" 'Is < Val(NotaMedia)
                                        MSFG_HstProvas.TextMatrix(lin, 7) = pgObsNH(RsMatriculaProva("ID"))
                                        MSFG_HstProvas.Row = MSFG_HstProvas.Rows - 1
                                        'MSFG_HstProvas.RowSel = MSFG_HstProvas.Rows - 1
                                        MSFG_HstProvas.Col = 0
                                        MSFG_HstProvas.ColSel = MSFG_HstProvas.Cols - 1
                                        MSFG_HstProvas.FillStyle = flexFillRepeat
                                        MSFG_HstProvas.CellForeColor = QBColor(12)
                                    Case "NC"
                                        MSFG_HstProvas.Row = MSFG_HstProvas.Rows - 1
                                        'MSFG_HstProvas.RowSel = MSFG_HstProvas.Rows - 1
                                        MSFG_HstProvas.Col = 0
                                        MSFG_HstProvas.ColSel = MSFG_HstProvas.Cols - 1
                                        MSFG_HstProvas.FillStyle = flexFillRepeat
                                        MSFG_HstProvas.CellForeColor = QBColor(0)
                                    Case Else
                                        MSFG_HstProvas.Row = MSFG_HstProvas.Rows - 1
                                        'MSFG_HstProvas.RowSel = MSFG_HstProvas.Rows - 1
                                        MSFG_HstProvas.Col = 0
                                        MSFG_HstProvas.ColSel = MSFG_HstProvas.Cols - 1
                                        MSFG_HstProvas.FillStyle = flexFillRepeat
                                        MSFG_HstProvas.CellForeColor = QBColor(0)
                                End Select
                        End If

                        RsProvas.MoveNext

                        'MSFG_HstProvas.Row = 1
                        'MSFG_HstProvas.Col = 1
                        'MSFG_HstProvas.RowSel = lin
                        'MSFG_HstProvas.ColSel = MSFG_HstProvas.Cols - 1
                        'MSFG_HstProvas.Sort = 1
                        lin = lin + 1
                        MSFG_HstProvas.Rows = MSFG_HstProvas.Rows + 1
                    Loop
                    'RsMatriculaSerie.MoveNext
            End If
            'RsTrafego.MoveNext
        'Loop
        'MSFG_HstProvas.Rows = MSFG_HstProvas.Rows - 1
        MSFG_HstProvas.Row = 0
        MSFG_HstProvas.Col = 0
        RsMatriculaSerie.MoveNext
    Loop
     'Alinha o Titulo das provas
    With MSFG_HstProvas
        
        If .Rows = 1 Then
            .Rows = 2
        End If
        .Col = 2
        .ColSel = 2
        .Row = 1
        .RowSel = .Rows - 1
        .FillStyle = flexFillRepeat
        .CellAlignment = 1
    End With
    'Organiza as provas
    With MSFG_HstProvas
        .Rows = .Rows - 1
        .Row = 1
        .RowSel = .Rows - 1
        .Col = 1
        .ColSel = .Cols - 1
        .FillStyle = flexFillRepeat
        .Sort = 1
    End With
    Exit Sub
TrtErroPrv:
    MsgBox Err.Description, vbCritical, Err.Number
    Resume Next
End Sub
Private Sub Cb_DisciplinaProvas_DropDown()
    MebMatricula.PromptInclude = False
    If MebMatricula.Text = "" Or Cb_EnsinoProvas.Text = "" Then
            MebMatricula.PromptInclude = True
            Exit Sub
        Else
            MebMatricula.PromptInclude = True
    End If
    Cb_DisciplinaProvas.Clear
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID) '& " AND isnull(DtConclusao)")
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.BOF Then
            MsgBox "Não existe nenhuma Disciplina cadastrada para esta Matricula.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsMatriculaDisciplina.MoveFirst
            Do Until RsMatriculaDisciplina.EOF
                'Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina WHERE DisciplinaID =" & RsMatriculaDisciplina.Fields("DisciplinaID"))
                'RsDisciplina.MoveFirst
                Disciplina = PgNomeDisciplina(RsMatriculaDisciplina.Fields("DisciplinaID")) 'RsDisciplina.Fields("Descr")
                Cb_DisciplinaProvas.AddItem (Disciplina)
                RsMatriculaDisciplina.MoveNext
            Loop
    End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Acao = 0
    'MatrID = ""
    SST_Matricula.Tab = 0
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas ORDER BY Nome")

End Sub
Private Sub Cb_Nome_Click()
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE Nome = '" & rc(Cb_Nome.Text) & "'")
    With RsMatricula
        If .BOF And .EOF Then
            MsgBox "Matricula não encontrada.", vbInformation, "CESNet - Aviso!"
            MebMatricula.SetFocus
            Exit Sub

        End If
        If Trim(.Fields("MatrID")) = "" Or IsNull(.Fields("MatrID")) Then
            Exit Sub
        End If
        MatrID = .Fields("MatrID")
        MebMatricula.Text = MatrID
    End With
   
    MstDadosAluno
    
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
 
    MatrID = MebMatricula.Text
    If Not IsNumeric(MatrID) Then
        LimpDados
        LimpaMatricula
        LimpConclDiscipl
        Exit Sub
    End If
    Cb_Nome.Text = PgDadosMatr(MatrID).Nome
        Lb_AtivoInativo.Caption = PgStatusMatricula(MatrID)
        If Lb_AtivoInativo.Caption = "ATIVO" Then
                Lb_AtivoInativo.ForeColor = vbBlue
            Else
                Lb_AtivoInativo.ForeColor = vbRed
        End If
        Meb_DtRetorno.PromptInclude = False
        Meb_DtRetorno.Text = PgDadosMatr(MatrID).DtRetorno
        Meb_DtRetorno.PromptInclude = True
        Cb_Nome.Text = PgDadosMatr(MatrID).Nome
        lbUnidade.Caption = PgDadosMatr(MatrID).UnidMatr
        DTP_ValidadeCard.Value = PgDadosMatr(MatrID).ValCard
        Txt_NumAntGeral.Text = PgDadosMatr(MatrID).NumAnt
        txtNumConexao.Text = PgDadosMatr(MatrID).NumConexao
        txtNumCenso.Text = PgDadosMatr(MatrID).NumCenso

    Cb_EnsinoDisciplinas.Clear
    MSFG_HstDisciplinas.Rows = 1
    MSFG_HstDisciplinas.Rows = 2
    Cb_EnsinoProvas.Clear
    Cb_DisciplinaProvas.Clear
    MSFG_HstProvas.Rows = 1
    MSFG_HstProvas.Rows = 2
    MSFG_EnsinoGeral.Rows = 1
    MSFG_EnsinoGeral.Rows = 2
    SST_Matricula.Tab = 0
    SST_Matricula_Click (0)

    '***** Checar Aviso ******
    If PgAviso(MatrID) = True Then
        Exit Sub
    End If
    '*************************
    
End Sub
Private Sub LimpDados()
    MebMatricula.PromptInclude = False
    
    MebMatricula.Text = ""
    Lb_AtivoInativo.Caption = ""
    Cb_Nome.Text = ""
    lbUnidade.Caption = ""
    
 
    

    
    MebMatricula.PromptInclude = True
 

End Sub


Private Sub Lb_SelTodas_Click()
    Dim i As Integer
    For i = 0 To Lst_Serie.ListCount - 1
        Lst_Serie.Selected(i) = True
    Next
End Sub

Private Sub Lb_SelTodasDisc00_Click()
    Dim i As Integer
    For i = 0 To Lst_Disciplina00.ListCount - 1
        Lst_Disciplina00.Selected(i) = True
    Next

End Sub

Private Sub Lb_SelTodasSerie00_Click()
    Dim i As Integer
    For i = 0 To Lst_Serie00.ListCount - 1
        Lst_Serie00.Selected(i) = True
    Next

End Sub

Private Sub Lst_Disciplina_Click()
    Bt_AlterarMatr.Caption = "Alterar Matricula em " & Lst_Disciplina.Text
    LstSeries
End Sub



Private Sub Meb_DtConclusao_GotFocus()
    Meb_DtConclusao.SelStart = 0
    Meb_DtConclusao.SelLength = 10
End Sub


Private Sub Meb_DtConclusao_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub

Private Sub MebMatricula_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 114 Then
       CarregarMatricula (formBuscar.IniciarBusca("Matriculas"))
    End If

End Sub

Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        'MatrID = Trim(MebMatricula.Text)
        MstDadosAluno
    End If
    
End Sub

Private Sub MSFG_EnsinoGeral_Click()
    Dim RsMatrEns As Recordset
    Dim CursoID As Integer
 
    
    If MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 1) <> "" And MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 2) = "" Then
            Bt_TrancarCursoGeral.Caption = "Trancar Curso"
            DTP_TrancaGeral.Enabled = True
            Txt_MotivoGeral.Enabled = True
            DTP_TrancaGeral.Value = IIf(MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 2) = "", Date, MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 2))
            Txt_MotivoGeral.Text = MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 3)
            Bt_TrancarCursoGeral.Enabled = True
        Else
            DTP_TrancaGeral.Enabled = False
            Txt_MotivoGeral.Enabled = False
            Bt_TrancarCursoGeral.Enabled = False
            CursoID = PgIDEnsino(MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 0))
            Set RsMatrEns = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & CursoID)
            If RsMatrEns.BOF And RsMatrEns.EOF Then
                    'MsgBox "Erro ao localizar Curso", vbInformation, "CESNet - Aviso"
                    RsMatrEns.Close
                    Exit Sub
                Else
                    RsMatrEns.MoveFirst
                    If RsMatrEns.Fields("Trancado") = True Then
                        DTP_TrancaGeral.Value = MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 2)
                        Txt_MotivoGeral.Text = MSFG_EnsinoGeral.TextMatrix(MSFG_EnsinoGeral.Row, 3)
                        Bt_TrancarCursoGeral.Caption = "Ativar Curso"
                        Bt_TrancarCursoGeral.Enabled = True
                    End If
                    RsMatrEns.Close
            End If
            
    End If
End Sub

Private Sub MSFG_HstDisciplinas_Click()
   If Trim(MSFG_HstDisciplinas.TextMatrix(MSFG_HstDisciplinas.Row, 1)) = "" Then
        btExcluirDisciplina.Enabled = False
        Exit Sub
    End If
    btExcluirDisciplina.Enabled = True
End Sub

Private Sub SST_Matricula_Click(PreviousTab As Integer)
    Select Case SST_Matricula.Tab
        Case 0 'Geral
            MSFG_EnsinoGeral.Rows = 1
            sstCurso.Tab = 0
            Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
            If RsEnsino.BOF And RsEnsino.EOF Then
                    Exit Sub
                Else
                    RsEnsino.MoveFirst
                    Do Until RsEnsino.EOF
                        Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & RsEnsino.Fields("ID"))
                        If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                            Else
                                With MSFG_EnsinoGeral
                                    MSFG_EnsinoGeral.Rows = MSFG_EnsinoGeral.Rows + 1
                                    .TextMatrix(.Rows - 1, 0) = RsEnsino.Fields("Descr")
                                    .TextMatrix(.Rows - 1, 1) = IIf(IsNull(RsMatriculaEnsino.Fields("DtInicio")), "", RsMatriculaEnsino.Fields("DtInicio"))
                                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsMatriculaEnsino.Fields("DtFinal")), "", RsMatriculaEnsino.Fields("DtFinal"))
                                    .TextMatrix(.Rows - 1, 3) = IIf(IsNull(RsMatriculaEnsino.Fields("Local")), "", RsMatriculaEnsino.Fields("Local"))
                                    .TextMatrix(.Rows - 1, 4) = IIf(Trim(.TextMatrix(.Rows - 1, 2)) = "", ChkDtProxRenovacao(MatrID, .TextMatrix(.Rows - 1, 1)), "00/00/0000")
                                    
                                    RsMatriculaEnsino.Close
                                End With
                                
                        End If
                        RsEnsino.MoveNext
                    Loop
                    'MSFG_EnsinoGeral.Rows = MSFG_EnsinoGeral.Rows - 1
            End If
            
        Case 1 'Matricula
            LimpaMatricula
            Cb_Ensino.Clear
            '***************** CHECAR *****************************************
            'checa se o aluno possui ensino concluido
            'Obs: melhorar checagem, inv. IF
            'If PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)) <> 0 Then
            '******************************************************************
            Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & Trim(MebMatricula.Text) & "'")
            If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                    Frame_Matr00.Visible = True
                    Frame_Matr01.Visible = False
                    dtpInicio00.Value = Date
                    Cb_Ensino00.Clear
                    Lst_Disciplina00.Clear
                    Lst_Serie00.Clear
                Else
                    Frame_Matr01.Visible = True
                    Frame_Matr00.Visible = False
                    dtpInicio.Value = IIf(IsNull(RsMatriculaEnsino.Fields("DtInicio")), Date, RsMatriculaEnsino.Fields("DtInicio"))
                    dtpInicio.Enabled = False
                    Cb_Ensino.AddItem (IIf(PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)) = 0, " ", PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False))))
                    Cb_Ensino.Text = IIf(PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)) = 0, " ", PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)))
            End If
        Case 2 'Disciplina
            Cb_EnsinoDisciplinas.Clear
            MSFG_HstDisciplinas.Rows = 1
            MSFG_HstDisciplinas.Rows = 2
            If PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)) <> 0 Then
                Cb_EnsinoDisciplinas.AddItem (PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)))
                Cb_EnsinoDisciplinas.Text = PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False))
            End If
            
            LimpaMatricula
        Case 3 'Provas
            Cb_EnsinoProvas.Clear
            If PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)) <> 0 Then
                Cb_EnsinoProvas.AddItem (PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False)))
                Cb_EnsinoProvas.Text = PgNomeEnsino(PgMatrEnsino(MebMatricula.Text, False))
            End If
            'Cb_DisciplinaProvas.Clear
            Cb_DisciplinaProvas.Clear
            MSFG_HstProvas.Rows = 1
            MSFG_HstProvas.Rows = 2
    End Select
End Sub

Private Sub Txt_CidadeConclusao_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If

End Sub



Private Sub Txt_MotivoGeral_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))

End Sub



Private Sub LstDisciplinas()
    If Cb_Ensino.Text = "" Then
        Exit Sub
    End If
    EnsinoID = PgIDEnsino(Cb_Ensino.Text)
    
    Set RsGradeEnsinoDisciplinas = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID)
    If RsGradeEnsinoDisciplinas.BOF And RsGradeEnsinoDisciplinas.EOF Then
            Lst_Disciplina.Clear
            Lst_Serie.Clear
            Exit Sub
        Else
            Lst_Disciplina.Clear
            RsGradeEnsinoDisciplinas.MoveFirst
            
            Do Until RsGradeEnsinoDisciplinas.EOF
                'Pegar o nome da Disciplina
                DisciplinaID = RsGradeEnsinoDisciplinas.Fields("DisciplinaID")
                Lst_Disciplina.AddItem (PgNomeDisciplina(RsGradeEnsinoDisciplinas.Fields("DisciplinaID"))) 'RsDisciplina.Fields("Descr"))
                RsGradeEnsinoDisciplinas.MoveNext
                Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                    Else
                        If IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) Or IsNull(RsMatriculaDisciplina.Fields("Local")) Then
                            Else
                        End If
                End If
            Loop
    End If
End Sub
Private Sub LstDisciplinas00()
    If Cb_Ensino00.Text = "" Then
        Exit Sub
    End If
    EnsinoID = PgIDEnsino(Cb_Ensino00.Text)
    
    Set RsGradeEnsinoDisciplinas = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID)
    If RsGradeEnsinoDisciplinas.BOF And RsGradeEnsinoDisciplinas.EOF Then
            Lst_Disciplina00.Clear
            Lst_Serie.Clear
            Exit Sub
        Else
            Lst_Disciplina00.Clear
            Lst_Serie00.Clear
            RsGradeEnsinoDisciplinas.MoveFirst
            
            Do Until RsGradeEnsinoDisciplinas.EOF
                'Pegar o nome da Disciplina
                DisciplinaID = RsGradeEnsinoDisciplinas.Fields("DisciplinaID")
                Lst_Disciplina00.AddItem (PgNomeDisciplina(RsGradeEnsinoDisciplinas.Fields("DisciplinaID"))) 'RsDisciplina.Fields("Descr"))
                RsGradeEnsinoDisciplinas.MoveNext
                Set RsGradeEnsinoSerie = BD.OpenRecordset("SELECT * FROM GradeEnsinoSeries WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                If RsGradeEnsinoSerie.BOF And RsGradeEnsinoSerie.EOF Then
                    Else
                        'Lst_Serie00.Clear
                        RsGradeEnsinoSerie.MoveFirst
                        Lst_Serie00.Enabled = False
                        Do Until RsGradeEnsinoSerie.EOF
                            
                            If LocSerie(PgNomeSerie(RsGradeEnsinoSerie.Fields("SerieID"))) = 0 Then
                                Lst_Serie00.AddItem (PgNomeSerie(RsGradeEnsinoSerie.Fields("SerieID")))
                            End If
                            
                            
                            RsGradeEnsinoSerie.MoveNext
                        Loop
                End If
            Loop
    End If
End Sub
Private Function LocSerie(txtSerie As String) As Integer
    Dim TmpSerie As Integer
    For TmpSerie = 0 To Lst_Serie00.ListCount - 1
        If Lst_Serie00.List(TmpSerie) = txtSerie Then
                LocSerie = 1
                Exit Function
            Else
                LocSerie = 0
        End If
    Next
End Function

Private Sub LstSeries()
    DisciplinaID = PgIDDisciplina(Lst_Disciplina.Text) 'RsDisciplina.Fields("DisciplinaID")
    Set RsGradeEnsinoSerie = BD.OpenRecordset("SELECT * FROM GradeEnsinoSeries WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY SerieID")
    If RsGradeEnsinoSerie.BOF And RsGradeEnsinoSerie.EOF Then
            Lst_Serie.Clear
            Exit Sub
        Else
            Lst_Serie.Clear
            RsGradeEnsinoSerie.MoveFirst
            Do Until RsGradeEnsinoSerie.EOF
                Set RsSerie = BD.OpenRecordset("SELECT * FROM Serie WHERE ID = " & RsGradeEnsinoSerie.Fields("SerieID"))
                Lst_Serie.AddItem (RsSerie.Fields("Descr"))
                SerieID = RsSerie.Fields("ID")
                RsGradeEnsinoSerie.MoveNext
                Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID & " ORDER BY SerieID")
                If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                    Else
                        Lst_Serie.Selected(Lst_Serie.ListCount - 1) = True
                End If
            Loop
    End If
    ChkDisciplina
End Sub
Private Sub ChkDisciplina()
        DoEvents
        Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
        If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            'Meb_DtConclusao.PromptInclude = False
            'Meb_DtConclusao.Text = ""
            'Cb_LocConclusao.Clear
            'Txt_CidadeConclusao.Text = ""
            'Cb_UFConclusao.Text = ""
            'Meb_DtConclusao.PromptInclude = True
            'Meb_DtConclusao.Enabled = True
            'Cb_LocConclusao.Enabled = True
            'Txt_CidadeConclusao.Enabled = True
            'Cb_UFConclusao.Enabled = False
            Call LimpConclDiscipl
            Call HDConclDiscipl(True)
            Bt_GrvConclusao.Enabled = True
        Else
            RsMatriculaDisciplina.MoveFirst
            If IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) Then ' And IsNull(RsMatriculaDisciplina.Fields("Local")) Then
                    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY SerieID") '& " AND SerieID = " & SerieID & " ORDER BY SerieID")
                    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                            'Meb_DtConclusao.PromptInclude = False
                            'Meb_DtConclusao.Text = ""
                            'Cb_LocConclusao.Clear
                            'Txt_CidadeConclusao.Text = ""
                            'Cb_UFConclusao.Text = ""
                            'Txt_CidadeConclusao.Enabled = True
                            'Cb_UFConclusao.Enabled = True
                            'Meb_DtConclusao.PromptInclude = True
                            'Meb_DtConclusao.Enabled = True
                            'Cb_LocConclusao.Enabled = True
                            Call LimpConclDiscipl
                            Call HDConclDiscipl(True)
                            Bt_GrvConclusao.Enabled = True
                        Else
                            Call LimpConclDiscipl
                            Call HDConclDiscipl(False)
                            Bt_GrvConclusao.Enabled = False
                    End If
                Else
                    RsMatriculaDisciplina.MoveFirst
                    Meb_DtConclusao.PromptInclude = False
                    Meb_DtConclusao.Text = IIf(IsNull(RsMatriculaDisciplina.Fields("DtConclusao")), "", RsMatriculaDisciplina.Fields("DtConclusao")) 'Right(RsMatriculaDisciplina.Fields("DtConclusao"), 7))
                    Cb_LocConclusao.AddItem (IIf(IsNull(RsMatriculaDisciplina.Fields("Local")), " ", RsMatriculaDisciplina.Fields("Local")))
                    Cb_LocConclusao.Text = IIf(IsNull(RsMatriculaDisciplina.Fields("Local")), " ", RsMatriculaDisciplina.Fields("Local"))
                    Txt_CidadeConclusao.Text = IIf(IsNull(RsMatriculaDisciplina.Fields("Cidade")), "", RsMatriculaDisciplina.Fields("Cidade"))
                    Cb_UFConclusao.Text = IIf(IsNull(RsMatriculaDisciplina.Fields("UF")), " ", RsMatriculaDisciplina.Fields("UF"))
                    
                    Meb_DtConclusao.PromptInclude = True
                    'Meb_DtConclusao.Enabled = True
                    'Cb_LocConclusao.Enabled = True
                    Call HDConclDiscipl(True)
                    Bt_GrvConclusao.Enabled = True
            End If
    End If
End Sub
Private Sub LimpaMatricula()
    Cb_Ensino.Clear
    Lst_Disciplina.Clear
    Lst_Serie.Clear
    dtpInicio00.Value = Date
    dtpInicio.Value = Date
    MSFG_EnsinoGeral.Rows = 1
    
    Txt_NumAntGeral.Text = ""
    Txt_MotivoGeral.Text = ""
    txtNumConexao.Text = ""
    txtNumCenso.Text = ""
    'Meb_DtConclusao.PromptInclude = False
    'Meb_DtConclusao.PromptInclude = False
    Call LimpConclDiscipl
    'Meb_DtConclusao.Text = ""
    'Cb_LocConclusao.Clear
    'Meb_DtConclusao.Enabled = False
    'Cb_LocConclusao.Enabled = False
    Bt_GrvConclusao.Enabled = False
    
    'Meb_DtConclusao.PromptInclude = True
    'Meb_DtConclusao.PromptInclude = True
    
End Sub


Private Sub GrvEnsino00()
'On Error GoTo TratErro:
    Dim cDisciplina As Integer
    Dim cSerie As Integer
   
    EnsinoID = PgIDEnsino(Trim(Cb_Ensino00.Text))
    
    'Checar Disciplina
    For cDisciplina = 0 To Lst_Disciplina00.ListCount - 1
        If Lst_Disciplina00.Selected(cDisciplina) = True Then
            'Gravar Ensino
            Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' and EnsinoID = " & EnsinoID)
            If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                RsMatriculaEnsino.AddNew
                RsMatriculaEnsino.Fields("DtInicio") = dtpInicio00.Value
                RsMatriculaEnsino.Fields("MatrID") = MatrID
                RsMatriculaEnsino.Fields("EnsinoID") = EnsinoID
                RsMatriculaEnsino.Fields("DtRenovacao") = ChkDtProxRenovacao(MatrID, dtpInicio00.Value)
                RsMatriculaEnsino.Fields("UsuarioID") = UsuarioID
                RsMatriculaEnsino.Fields("DtHrSis") = Now()
                RsMatriculaEnsino.Update
            End If
            'Gravar Disciplina
            DisciplinaID = PgIDDisciplina(Lst_Disciplina00.List(cDisciplina))
            Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' and EnsinoID = " & EnsinoID & " and DisciplinaID = " & DisciplinaID)
            If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                RsMatriculaDisciplina.AddNew
                RsMatriculaDisciplina.Fields("MatrID") = MatrID
                RsMatriculaDisciplina.Fields("EnsinoID") = EnsinoID
                RsMatriculaDisciplina.Fields("DisciplinaID") = DisciplinaID
                RsMatriculaDisciplina.Fields("UsuarioID") = UsuarioID
                RsMatriculaDisciplina.Fields("DtHrSis") = Now()
                RsMatriculaDisciplina.Update
            End If
            'Checar Serie
            For cSerie = 0 To Lst_Serie00.ListCount - 1
                If Lst_Serie00.Selected(cSerie) = True Then
                    SerieID = PgIDSerie(Lst_Serie00.List(cSerie))
                    'Checa se a serie é de competencia da disciplina
                    Set RsGradeEnsinoSerie = BD.OpenRecordset("SELECT * FROM GradeEnsinoSeries WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SERIEID = " & SerieID)
                    If RsGradeEnsinoSerie.BOF And RsGradeEnsinoSerie.EOF Then
                        Else
                            Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
                            If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                                RsMatriculaSerie.AddNew
                                RsMatriculaSerie.Fields("MatrID") = MatrID
                                RsMatriculaSerie.Fields("EnsinoID") = EnsinoID
                                RsMatriculaSerie.Fields("DisciplinaID") = DisciplinaID
                                RsMatriculaSerie.Fields("SerieID") = SerieID
                                RsMatriculaSerie.Fields("UsuarioID") = UsuarioID
                                RsMatriculaSerie.Fields("DtHrSis") = Now()
                                RsMatriculaSerie.Update
                            End If
                    End If
                End If
            Next
        End If
    Next
    Exit Sub
TratErro:
    Call RegLogErros(Err.Number, "Erro ao incluir as diciplinas em Matr00. - " & Err.Description, Me.Caption, UsuarioID)
    MsgBox "Erro ao cadastrar as diciplinas." & Chr(13) & "Descrição: " & Err.Description, vbInformation, Err.Number
    Resume Next
End Sub
Private Sub GrvSeries()
    For lin = 0 To Lst_Serie.ListCount - 1
        SerieID = PgIDSerie(Lst_Serie.List(lin))
        Select Case Lst_Serie.Selected(lin)
            Case True 'Caixa marcada
                Bt_GrvConclusao.Enabled = False
                Call HDConclDiscipl(False)
                'GRAVA ENSINO
                If PgMatrEnsino(MatrID, False) <> 0 Then
                        If PgMatrEnsino(MatrID, False) = EnsinoID Then
                                Else
                                    MsgBox "Esta matrícula já possue o curso " & PgNomeEnsino(PgMatrEnsino(MatrID, False)) & ", iniciado.", vbInformation, "CESNet - Aviso!"
                                    'Lst_Serie.Selected(Lst_Serie.ListIndex) = False
                                    Lst_Serie.Clear
                                    Exit Sub
                        End If
                    Else
                        Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID =" & EnsinoID)
                        If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
                            RsMatriculaEnsino.AddNew
                            
                            RsMatriculaEnsino.Fields("MatrID") = MatrID
                            RsMatriculaEnsino.Fields("DtInicio") = dtpInicio.Value
                            RsMatriculaEnsino.Fields("EnsinoID") = EnsinoID
                            RsMatriculaEnsino.Fields("DtRenovacao") = ChkDtProxRenovacao(MatrID, dtpInicio.Value)
                            RsMatriculaEnsino.Fields("UsuarioID") = UsuarioID
                            RsMatriculaEnsino.Fields("DtHrSis") = Now()
                            RsMatriculaEnsino.Update
                        Else
                            If MsgBox("Curso já concluido! Deseja reativa-lo?", vbInformation + vbYesNo, "CESNet - Aviso") = vbYes Then
                                    RsMatriculaEnsino.MoveFirst
                                    RsMatriculaEnsino.Edit
                                    RsMatriculaEnsino.Fields("DtFinal") = Null
                                    RsMatriculaEnsino.Fields("Local") = Null
                                    RsMatriculaEnsino.Fields("OcorrenciaID") = Null
                                    RsMatriculaEnsino.Fields("DtOcorrencia") = Null
                                    RsMatriculaEnsino.Fields("UsuarioID") = UsuarioID
                                    RsMatriculaEnsino.Fields("DtHrSis") = Now()
                                    RsMatriculaEnsino.Update
                                Else
                                    RsMatriculaEnsino.Close
                                    Exit Sub
                            End If
                        End If
                End If
                                
                'GRAVAR Disciplina
                Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                        RsMatriculaDisciplina.AddNew
                        RsMatriculaDisciplina.Fields("MatrID") = MatrID
                        RsMatriculaDisciplina.Fields("EnsinoID") = EnsinoID
                        RsMatriculaDisciplina.Fields("DisciplinaID") = DisciplinaID
                        RsMatriculaDisciplina.Fields("UsuarioID") = UsuarioID
                        RsMatriculaDisciplina.Fields("DtHrSis") = Now()
                        RsMatriculaDisciplina.Update
                        Call RegLog(MatrID, "INCLUSAO DE DISCIPLINA - " & PgNomeEnsino(EnsinoID) & " / " & PgNomeDisciplina(DisciplinaID))
                    Else
                        'Caso a disciplina ja esteja concluida
                        RsMatriculaDisciplina.MoveFirst
                        If Not IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) Then
                            If MsgBox("Disciplina já concluida em " & RsMatriculaDisciplina.Fields("DtConclusao") & "." & _
                                        vbCrLf & "Deseja reativa-la?", vbInformation + vbYesNo, "CESNet - Reativação de Disciplina") = vbYes Then
                                        'vbCrLf & "A reativação apagará todas as provas!"
                                Call RegLog(MatrID, "Ativou disciplina concluida em " & RsMatriculaDisciplina.Fields("DtConclusao") & " na " & RsMatriculaDisciplina.Fields("Local"))
                                RsMatriculaDisciplina.Edit
                                RsMatriculaDisciplina.Fields("DtConclusao") = Null
                                RsMatriculaDisciplina.Fields("Local") = Null
                                RsMatriculaDisciplina.Fields("Cidade") = Null
                                RsMatriculaDisciplina.Fields("UF") = Null
                                RsMatriculaDisciplina.Fields("UsuarioID") = UsuarioID
                                RsMatriculaDisciplina.Fields("DtHrSis") = Now()
                                RsMatriculaDisciplina.Update
                                Call LimpConclDiscipl
                                'Apagar a data de conclusao de ensino
                                Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
                                If RsMatriculaEnsino.BOF And RsMatricula.EOF Then
                                    Else
                                        RsMatriculaEnsino.MoveFirst
                                        Call RegLog(MatrID, "Ativou Curso concluido em " & RsMatriculaEnsino.Fields("DtFinal") & " na " & RsMatriculaEnsino.Fields("Local"))
                                        RsMatriculaEnsino.Edit
                                        RsMatriculaEnsino.Fields("DtFinal") = Null
                                        RsMatriculaEnsino.Fields("Local") = Null
                                        RsMatriculaEnsino.Fields("UsuarioID") = UsuarioID
                                        RsMatriculaEnsino.Fields("DtHrSis") = Now()
                                        RsMatriculaEnsino.Update
                                End If
                                'Apaga as PROVAS e SERIES
                                Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
                                If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                                        RsMatriculaSerie.AddNew
                                        RsMatriculaSerie.Fields("MatrID") = MatrID
                                        RsMatriculaSerie.Fields("EnsinoID") = EnsinoID
                                        RsMatriculaSerie.Fields("DisciplinaID") = DisciplinaID
                                        RsMatriculaSerie.Fields("SerieID") = SerieID
                                        RsMatriculaSerie.Fields("UsuarioID") = UsuarioID
                                        RsMatriculaSerie.Fields("DtHrSis") = Now()
                                        RsMatriculaSerie.Update
                                    Else
                                        RsMatriculaSerie.MoveFirst
                                        'APAGA AS PROVAS CASO A DISCIPLINA SEJA REATIVADA
                                        'Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & RsMatriculaSerie.Fields("EnsinoID") & _
                                        '                " AND DisciplinaID = " & RsMatriculaSerie.Fields("DisciplinaID") & " AND SerieID = " & RsMatriculaSerie.Fields("SerieID"))
                                        'If RsTrafego.BOF And RsTrafego.EOF Then
                                        '    Else
                                        '        RsTrafego.MoveFirst
                                        '        Do Until RsTrafego.EOF
                                        '            'BD.Execute "DELETE * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciiplinaID = " & DisciplinaID & " AND SerieID = " & SerieID
                                        '            BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID")
                                        '            BD.Execute "DELETE * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID")
                                        '            RsTrafego.MoveNext
                                        '        Loop
                                        'End If
                                        RsMatriculaSerie.Edit
                                        RsMatriculaSerie.Fields("DtIni") = Null
                                        RsMatriculaSerie.Fields("DtFinal") = Null
                                        RsMatriculaSerie.Fields("Aprovado") = False
                                        RsMatriculaSerie.Fields("UsuarioID") = UsuarioID
                                        RsMatriculaSerie.Fields("DtHrSis") = Now()
                                        RsMatriculaSerie.Update
                                End If
            
                            End If
                        End If
                End If
                'GRAVAR SERIE
                If Meb_DtConclusao.Text = "" Or Cb_LocConclusao.Text = "" Then
                    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
                    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                        With RsMatriculaSerie
                            .AddNew
                            .Fields("MatrID") = MatrID
                            .Fields("EnsinoID") = EnsinoID
                            .Fields("DisciplinaID") = DisciplinaID
                            .Fields("SerieID") = SerieID
                            .Fields("UsuarioID") = UsuarioID
                            .Fields("DtHrSis") = Now()
                            .Update
                        End With
                    End If
                End If
            Case False 'Caixa Desmarcada
                Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
                If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
                    Else
                        RsMatriculaSerie.MoveFirst
                        Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & RsMatriculaSerie.Fields("EnsinoID") & _
                                    " AND DisciplinaID = " & RsMatriculaSerie.Fields("DisciplinaID") & " AND SerieID = " & RsMatriculaSerie.Fields("SerieID"))
                        If RsTrafego.BOF And RsTrafego.EOF Then
                            Else
                                RsTrafego.MoveFirst
                                Do Until RsTrafego.EOF
                                    'Grava no arquivo de log com as provas excluidas
                                    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID"))
                                    If RsTMP.BOF And RsTMP.EOF Then
                                            Call RegLog(MatrID, "EXCLUSAO DE DISCIPLINA (sem provas efetuadas) - " & PgNomeDisciplina(RsMatriculaSerie.Fields("DisciplinaID")) & " / " & PgNomeSerie(SerieID))
                                            RsTMP.Close
                                        Else
                                            RsTMP.MoveFirst
                                            Do Until RsTMP.EOF
                                                Call RegLog(MatrID, "EXCLUSAO DE PROVA - tela de Matricula - " & PgNomeEnsino(RsTMP.Fields("EnsinoID")) & "/" & PgNomeDisciplina(RsTMP.Fields("DisciplinaID")) & " - Prova: " & RsTMP.Fields("NProva") & " Avaliacao: " & RsTMP.Fields("DtAvaliacao") & " TIPO: " & RsTMP.Fields("Tipo") & " - Nota: " & RsTMP.Fields("Nota"))
                                                RsTMP.MoveNext
                                            Loop
                                            RsTMP.Close
                                    End If
                                    BD.Execute "DELETE * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID")
                                    BD.Execute "DELETE * FROM ProvasTMP WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RsTrafego.Fields("RefTrafegoID")
                                    RsTrafego.MoveNext
                                Loop
                        End If
                        
                        RsMatriculaSerie.Delete
                        
                End If
                tmp = 0
                Do Until tmp = Lst_Serie.ListCount
                    If Lst_Serie.Selected(tmp) = True Then
                        Call HDConclDiscipl(False)
                        Bt_GrvConclusao.Enabled = False
                        Exit Do
                    End If
                
                    Call HDConclDiscipl(True)
                    Bt_GrvConclusao.Enabled = True
                    tmp = tmp + 1
                Loop
                If Lst_Serie.ListCount = tmp Then
                    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE  MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                    If RsTMP.BOF And RsTMP.EOF Then
                        Else
                            RsTMP.MoveFirst
                            RsTMP.Delete
                            Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "'") ' AND EnsinoID = " & EnsinoID)
                            If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
                            Call RegLog(MatrID, "Excluido curso " & PgNomeEnsino(EnsinoID))
                                BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID
                            End If
                    End If
                End If
        End Select
    Next
    'Call RegLog(MatrID, "Alterou a Matricula/Cadastro.")
End Sub
Private Sub LimpConclDiscipl()
    Meb_DtConclusao.PromptInclude = False
    Meb_DtConclusao.Text = ""
    Meb_DtConclusao.PromptInclude = True
    Cb_LocConclusao.Clear
    Txt_CidadeConclusao.Text = ""
    Cb_UFConclusao.Text = " "
End Sub
Private Sub HDConclDiscipl(op As Boolean)
    Meb_DtConclusao.Enabled = op
    Cb_LocConclusao.Enabled = op
    Txt_CidadeConclusao.Enabled = op
    Cb_UFConclusao.Enabled = op
    Bt_GrvConclusao.Enabled = op
End Sub

Public Sub CarregarMatricula(tmpMatrID As String)
    If Trim(tmpMatrID) = "" Then Exit Sub
    MatrID = tmpMatrID
    Me.Show
    MebMatricula.PromptInclude = False
    MebMatricula.Text = MatrID
    MebMatricula.PromptInclude = True
    MstDadosAluno
End Sub
