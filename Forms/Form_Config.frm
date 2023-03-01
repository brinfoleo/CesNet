VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Form_Config 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Configurações do Aplicativo"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8700
   Icon            =   "Form_Config.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   8700
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   730
      Left            =   6420
      Picture         =   "Form_Config.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   6510
      Width           =   2130
   End
   Begin VB.CommandButton Bt_Gravar 
      Caption         =   "&Gravar"
      Height          =   730
      Left            =   4260
      Picture         =   "Form_Config.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6510
      Width           =   2115
   End
   Begin TabDlg.SSTab SST_Config 
      Height          =   6090
      Left            =   60
      TabIndex        =   1
      Top             =   360
      Width           =   8565
      _ExtentX        =   15108
      _ExtentY        =   10742
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      TabCaption(0)   =   "Geral"
      TabPicture(0)   =   "Form_Config.frx":091E
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Disciplinas"
      TabPicture(1)   =   "Form_Config.frx":093A
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "SSTab1"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      Begin TabDlg.SSTab SSTab1 
         Height          =   5475
         Left            =   180
         TabIndex        =   28
         Top             =   480
         Width           =   8175
         _ExtentX        =   14420
         _ExtentY        =   9657
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "Regras de Conclusão"
         TabPicture(0)   =   "Form_Config.frx":0956
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "Frame6"
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "Outras"
         TabPicture(1)   =   "Form_Config.frx":0972
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Frame2"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame4"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame5"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.Frame Frame5 
            Caption         =   "Módulos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1275
            Left            =   240
            TabIndex        =   47
            Top             =   1800
            Width           =   4275
            Begin VB.ComboBox Cb_VincMod 
               Height          =   315
               ItemData        =   "Form_Config.frx":098E
               Left            =   3000
               List            =   "Form_Config.frx":0998
               Style           =   2  'Dropdown List
               TabIndex        =   48
               Top             =   600
               Width           =   1080
            End
            Begin MSMask.MaskEdBox Meb_MaxModulos 
               Height          =   285
               Left            =   3000
               TabIndex        =   49
               Top             =   225
               Width           =   615
               _ExtentX        =   1085
               _ExtentY        =   503
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   2
               Mask            =   "##"
               PromptChar      =   "_"
            End
            Begin VB.Label Label5 
               Alignment       =   1  'Right Justify
               Caption         =   "Quantidade de disciplinas em curso:"
               Height          =   195
               Left            =   300
               TabIndex        =   51
               Top             =   270
               Width           =   2625
            End
            Begin VB.Label Label6 
               Alignment       =   1  'Right Justify
               Caption         =   "Vincular Avaliação com o Emprestimo de Modulo (s):"
               Height          =   435
               Left            =   180
               TabIndex        =   50
               Top             =   540
               Width           =   2745
            End
         End
         Begin VB.Frame Frame4 
            Caption         =   "Sistema de notas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1815
            Left            =   240
            TabIndex        =   42
            Top             =   3240
            Width           =   4275
            Begin VB.OptionButton Opt_SisNota 
               Caption         =   "Sintético - Informar somente se foi aprovado."
               Height          =   195
               Index           =   1
               Left            =   135
               TabIndex        =   44
               Top             =   720
               Width           =   3480
            End
            Begin VB.OptionButton Opt_SisNota 
               Caption         =   "Analítico - Informar o  percentual (%) de aprovação."
               Height          =   435
               Index           =   0
               Left            =   135
               TabIndex        =   43
               Top             =   240
               Width           =   3810
            End
            Begin VB.Label Label8 
               Caption         =   "Obs.:"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   400
                  Underline       =   -1  'True
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H00808080&
               Height          =   195
               Left            =   180
               TabIndex        =   46
               Top             =   1035
               Width           =   420
            End
            Begin VB.Label Label7 
               Caption         =   "Em caso sintético, o  CESNet aprovará o aluno usando o percentual(%) minimo para aprovação."
               ForeColor       =   &H00808080&
               Height          =   600
               Left            =   675
               TabIndex        =   45
               Top             =   1080
               Width           =   3345
            End
         End
         Begin VB.Frame Frame2 
            Caption         =   "Media de Aprovação:"
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
            Left            =   240
            TabIndex        =   39
            Top             =   540
            Width           =   4275
            Begin MSMask.MaskEdBox Meb_NotaMedia 
               Height          =   315
               Left            =   1785
               TabIndex        =   40
               Top             =   435
               Width           =   705
               _ExtentX        =   1244
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   4
               Mask            =   "####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label3 
               Alignment       =   1  'Right Justify
               Caption         =   "Media de Aprovação:"
               Height          =   210
               Left            =   150
               TabIndex        =   41
               Top             =   480
               Width           =   1560
            End
         End
         Begin VB.Frame Frame6 
            Caption         =   "Regras para conclusao de Curso"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   4935
            Left            =   -74820
            TabIndex        =   29
            Top             =   420
            Width           =   7815
            Begin VB.ListBox Lst_Disciplinas 
               Height          =   3435
               Left            =   120
               Style           =   1  'Checkbox
               TabIndex        =   35
               Top             =   900
               Width           =   3825
            End
            Begin VB.ComboBox Cb_Ensino 
               Height          =   315
               Left            =   750
               Style           =   2  'Dropdown List
               TabIndex        =   34
               Top             =   360
               Width           =   3195
            End
            Begin VB.CommandButton Bt_GrvDisciplinas 
               Caption         =   "Gravar regras de conclusão da grade"
               Height          =   1050
               Left            =   4080
               TabIndex        =   33
               Top             =   3780
               Width           =   3630
            End
            Begin VB.Frame Frame11 
               Caption         =   "Disciplinas Opcionais"
               BeginProperty Font 
                  Name            =   "MS Sans Serif"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1140
               Left            =   4140
               TabIndex        =   30
               Top             =   1500
               Width           =   3525
               Begin VB.TextBox Txt_NumMinDiscipl 
                  Height          =   285
                  Left            =   2640
                  MaxLength       =   1
                  TabIndex        =   31
                  Top             =   660
                  Width           =   735
               End
               Begin VB.Label Label11 
                  Alignment       =   1  'Right Justify
                  Caption         =   "Informe a quantidade mínima de disciplinas optativas a serem cursadas:"
                  Height          =   600
                  Left            =   180
                  TabIndex        =   32
                  Top             =   360
                  Width           =   2370
               End
            End
            Begin VB.Label Label12 
               Caption         =   "As disciplinas marcadas são obrigatórias."
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
               Left            =   120
               TabIndex        =   38
               Top             =   4560
               Width           =   3765
            End
            Begin VB.Label Label13 
               Caption         =   "Curso:"
               Height          =   195
               Left            =   225
               TabIndex        =   37
               Top             =   390
               Width           =   435
            End
            Begin VB.Label Label14 
               Caption         =   "Disciplinas:"
               Height          =   195
               Left            =   165
               TabIndex        =   36
               Top             =   675
               Width           =   825
            End
         End
      End
      Begin VB.Frame Frame1 
         Height          =   5595
         Left            =   -74880
         TabIndex        =   4
         Top             =   360
         Width           =   8355
         Begin VB.CheckBox chkBloqRenovVencida 
            Caption         =   "Bloquear matriculas não renovadas"
            Height          =   195
            Left            =   120
            TabIndex        =   56
            Top             =   4860
            Width           =   3795
         End
         Begin VB.Frame Frame7 
            Caption         =   "Historico Escolar"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   120
            TabIndex        =   52
            Top             =   3420
            Width           =   8115
            Begin VB.ComboBox cbonmDocHistEsc 
               Height          =   315
               Left            =   2040
               Style           =   2  'Dropdown List
               TabIndex        =   55
               Top             =   720
               Width           =   5655
            End
            Begin VB.CheckBox chkUsarHistEscImp 
               Caption         =   "Usar  Historico Escolar Importado ao CESNet com o Gerenciador de Declarações"
               Height          =   255
               Left            =   240
               TabIndex        =   53
               Top             =   300
               Width           =   7635
            End
            Begin VB.Label Label19 
               Alignment       =   1  'Right Justify
               Caption         =   "Selecione o Dcoumento:"
               Height          =   195
               Left            =   60
               TabIndex        =   54
               Top             =   780
               Width           =   1875
            End
         End
         Begin VB.CheckBox chkCartConjugada 
            Caption         =   "Imprimir carteira junto da ficha de cadastro"
            Height          =   195
            Left            =   4560
            TabIndex        =   27
            Top             =   5220
            Width           =   3315
         End
         Begin VB.Frame Frame8 
            Caption         =   "Encaminhamento a Orientação:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   4500
            TabIndex        =   23
            Top             =   600
            Width           =   3735
            Begin VB.TextBox Txt_NumRepro 
               Height          =   285
               Left            =   2760
               MaxLength       =   1
               TabIndex        =   24
               Top             =   360
               Width           =   735
            End
            Begin VB.Label Label15 
               Alignment       =   1  'Right Justify
               Caption         =   "Número de reprovações para encaminhamento ao prof. orientador:"
               Height          =   435
               Left            =   120
               TabIndex        =   26
               Top             =   240
               Width           =   2595
            End
            Begin VB.Label Label18 
               Caption         =   "Obs: Colocando ""0"" (zero) o aluno fica obrigado a passar na orientação antes da avaliação."
               ForeColor       =   &H00808080&
               Height          =   675
               Left            =   180
               TabIndex        =   25
               Top             =   720
               Width           =   3495
            End
         End
         Begin VB.Frame Frame10 
            Caption         =   "Resultado de Provas:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1515
            Left            =   120
            TabIndex        =   14
            Top             =   600
            Width           =   4275
            Begin VB.TextBox Txt_RPMax 
               Height          =   285
               Left            =   2790
               MaxLength       =   2
               TabIndex        =   17
               Top             =   630
               Width           =   870
            End
            Begin VB.TextBox Txt_RPTime 
               Height          =   285
               Left            =   2790
               MaxLength       =   3
               TabIndex        =   16
               Top             =   240
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.TextBox Txt_RPMaxReg 
               Height          =   285
               Left            =   2790
               MaxLength       =   3
               TabIndex        =   15
               Top             =   1035
               Visible         =   0   'False
               Width           =   870
            End
            Begin VB.Label Label10 
               Alignment       =   1  'Right Justify
               Caption         =   "Maximo de Resultados por tela:"
               Height          =   195
               Left            =   255
               TabIndex        =   20
               Top             =   660
               Width           =   2445
            End
            Begin VB.Label Label9 
               Alignment       =   1  'Right Justify
               Caption         =   "Tempo de Atualização de Tela (seg.):"
               Height          =   195
               Left            =   90
               TabIndex        =   19
               Top             =   270
               Visible         =   0   'False
               Width           =   2670
            End
            Begin VB.Label Label16 
               Alignment       =   1  'Right Justify
               Caption         =   "Maximo de Resultados por consulta:"
               Height          =   375
               Left            =   180
               TabIndex        =   18
               Top             =   945
               Visible         =   0   'False
               Width           =   2490
            End
         End
         Begin VB.CheckBox Chk_DeslWin 
            Caption         =   "Desligar o Computador ao encerrar o CESNet."
            Height          =   195
            Left            =   180
            TabIndex        =   13
            Top             =   5220
            Visible         =   0   'False
            Width           =   3585
         End
         Begin VB.ComboBox Cb_Unidade 
            Height          =   315
            Left            =   810
            TabIndex        =   12
            Text            =   "Cb_Unidade"
            Top             =   180
            Width           =   915
         End
         Begin VB.Frame Frame3 
            Caption         =   "Acesso ao sistema:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   795
            Left            =   4500
            TabIndex        =   9
            Top             =   2220
            Width           =   2655
            Begin MSMask.MaskEdBox Meb_MaxAcessos 
               Height          =   315
               Left            =   1650
               TabIndex        =   10
               Top             =   300
               Width           =   735
               _ExtentX        =   1296
               _ExtentY        =   556
               _Version        =   393216
               PromptInclude   =   0   'False
               MaxLength       =   5
               Mask            =   "#####"
               PromptChar      =   "_"
            End
            Begin VB.Label Label1 
               Alignment       =   1  'Right Justify
               Caption         =   "Maximo de acessos:"
               Height          =   195
               Left            =   135
               TabIndex        =   11
               Top             =   360
               Width           =   1440
            End
         End
         Begin VB.Frame Frame9 
            Caption         =   "Aluno Inativos:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1095
            Left            =   120
            TabIndex        =   5
            Top             =   2220
            Width           =   3735
            Begin VB.TextBox Txt_InatDias 
               Height          =   285
               Left            =   2655
               MaxLength       =   4
               TabIndex        =   7
               Top             =   315
               Width           =   870
            End
            Begin VB.CheckBox Chk_InatBloc 
               Caption         =   "Bloquear matrícula inativa."
               Height          =   195
               Left            =   360
               TabIndex        =   6
               Top             =   765
               Width           =   2355
            End
            Begin VB.Label Label17 
               Alignment       =   1  'Right Justify
               Caption         =   "Número de dias sem atividades:"
               Height          =   195
               Left            =   135
               TabIndex        =   8
               Top             =   360
               Width           =   2490
            End
         End
         Begin VB.Label Lb_Unid 
            Caption         =   "<Nome da Unidade>"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00FF0000&
            Height          =   255
            Left            =   1860
            TabIndex        =   22
            Top             =   225
            Width           =   6420
         End
         Begin VB.Label Label4 
            Caption         =   "Unidade:"
            Height          =   195
            Left            =   90
            TabIndex        =   21
            Top             =   240
            Width           =   675
         End
      End
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CONFIGURAÇÕES"
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
      Top             =   0
      Width           =   8715
   End
End
Attribute VB_Name = "Form_Config"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsEnsino                   As Recordset
Dim RsConfig                   As Recordset
Dim RsGradeEnsinoDisciplina    As Recordset
Dim RsUnid                     As Recordset

Dim EnsinoID    As Integer
Dim DisciplinaID    As Integer

Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub

Private Sub Bt_GrvDisciplinas_Click()
    Dim tmp As Integer
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    EnsinoID = PgIDEnsino(Cb_Ensino.Text)
    
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino WHERE ID = " & EnsinoID)
    If RsEnsino.BOF And RsEnsino.EOF Then
            MsgBox "Erro ao localizar Ensino na tabEnsino"
            Exit Sub
        Else
            RsEnsino.MoveFirst
            RsEnsino.Edit
            RsEnsino.Fields("NumMinDiscipl") = IIf(Trim(Txt_NumMinDiscipl.Text) = "", 0, Trim(Txt_NumMinDiscipl.Text))
            RsEnsino.Update
    End If
    Set RsGradeEnsinoDisciplina = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID & " ORDER BY DisciplinaID")
    For tmp = 0 To Lst_Disciplinas.ListCount - 1
        DisciplinaID = PgIDDisciplina(Lst_Disciplinas.List(tmp))
        If Lst_Disciplinas.Selected(tmp) = True Then
                Set RsGradeEnsinoDisciplina = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                RsGradeEnsinoDisciplina.Edit
                RsGradeEnsinoDisciplina.Fields("Obrigatoria") = True
                RsGradeEnsinoDisciplina.Update
            Else
                Set RsGradeEnsinoDisciplina = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                RsGradeEnsinoDisciplina.Edit
                RsGradeEnsinoDisciplina.Fields("Obrigatoria") = False
                RsGradeEnsinoDisciplina.Update
        End If
    Next
    MsgBox "Regras de conclusão de disciplinas alteradas com sucesso!", vbInformation, "CESNet - Aviso"
End Sub

Private Sub Cb_Ensino_Click()
    EnsinoID = PgIDEnsino(Cb_Ensino.Text)
    Txt_NumMinDiscipl.Text = pgNumMinDiscipl(EnsinoID)
    Set RsGradeEnsinoDisciplina = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID & " ORDER BY DisciplinaID")
    With RsGradeEnsinoDisciplina
        If .BOF And .EOF Then
                Lst_Disciplinas.Clear
            Else
                Lst_Disciplinas.Clear
                .MoveFirst
                Do Until .EOF
                    Lst_Disciplinas.AddItem (PgNomeDisciplina(.Fields("DisciplinaID")))
                    If .Fields("Obrigatoria") = True Then
                        Lst_Disciplinas.Selected(Lst_Disciplinas.ListCount - 1) = True
                    End If
                    .MoveNext
                Loop
        End If
    End With
End Sub

Private Sub Cb_Ensino_DropDown()
    Cb_Ensino.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.EOF Then
            Cb_Ensino.Clear
            Lst_Disciplinas.Clear
            Exit Sub
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_Ensino.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If
End Sub

Private Sub Cb_Unidade_Click()
    'Cb_Unidade.Text = RsConfig.Fields("Unidade")
    RsUnid.FindFirst "UnidID = '" & left("000", 3 - Len(Trim(CB_Unidade.Text))) & Trim(CB_Unidade.Text) & "'"
    If RsUnid.NoMatch Then
            MsgBox "Erro na procura pela Unidade de Ensino. Por favor Verifique", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            Lb_Unid.Caption = RsUnid.Fields("Nome")
    End If
End Sub

Private Sub Cb_Unidade_KeyPress(KeyAscii As Integer)
    KeyAscii = 0
End Sub

Private Sub Bt_Gravar_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    If Meb_MaxAcessos.Text = "" Or Meb_MaxAcessos.Text = 0 Then
        MsgBox "O campo Maximo de Acessos nao pode ser Vazio ou Zerado", vbInformation, "CESNet - Aviso!"
        Exit Sub
    End If
    
    RsConfig.Edit
    RsConfig.Fields("Unidade") = CB_Unidade.Text
    'RsConfig.Fields("LocalBD") = Txt_LocalBD.Text
    RsConfig.Fields("MaxAcessos") = Meb_MaxAcessos.Text
    RsConfig.Fields("NotaMedia") = Meb_NotaMedia.Text
    RsConfig.Fields("MaxModulosEmpr") = Meb_MaxModulos.Text
    RsConfig.Fields("CartConjugada") = chkCartConjugada.Value
    RsConfig.Fields("VincModulos") = IIf(Cb_VincMod.Text = "SIM", True, False)
    RsConfig.Fields("SisNota") = IIf(Opt_SisNota(0).Value = True, True, False)
    'RsConfig.Fields("pgNumMinDiscipl") = IIf(Trim(Txt_pgNumMinDiscipl.Text) = "", 0, Trim(Txt_pgNumMinDiscipl.Text))
    RsConfig.Fields("RPTime") = IIf(Trim(Txt_RPTime.Text) = "" Or Trim(Txt_RPTime.Text) = "0", 1, Trim(Txt_RPTime.Text))
    RsConfig.Fields("RPMax") = IIf(Trim(Txt_RPMax.Text) = "", 1, Trim(Txt_RPMax.Text))
    RsConfig.Fields("RPMaxReg") = IIf(Trim(Txt_RPMaxReg.Text) = "", 1, Trim(Txt_RPMaxReg.Text))
    
    RsConfig.Fields("NumRepro") = IIf(Trim(Txt_NumRepro.Text) = "", 0, Trim(Txt_NumRepro.Text))
    
    RsConfig.Fields("InatDias") = IIf(Trim(Txt_InatDias.Text) = "" Or Trim(Txt_InatDias.Text) < 1, 1, Trim(Txt_InatDias.Text))
    RsConfig.Fields("BloquearInativo") = Chk_InatBloc.Value
    
    RsConfig.Fields("BloqRenovVencida") = chkBloqRenovVencida.Value
    
    '////////////////////////////////////////////////////////////////////////
     
    RsConfig.Fields("UsarHistEscImp") = chkUsarHistEscImp.Value
    RsConfig.Fields("nmDocHistEsc") = IIf(Trim(cbonmDocHistEsc.Text) = "", Null, PgNomeDoc(Trim(cbonmDocHistEsc.Text)))
    
    '//////////////////////////////////////////////
    
    RsConfig.Fields("DeslWin") = IIf(Chk_DeslWin.Value = 1, True, False)
    RsConfig.Update
    'NotaMedia = Meb_NotaMedia.Text
    'DeslWin = IIf(Chk_DeslWin.Value = 1, True, False)
    Call PgRegrasSis
    Unload Me
End Sub

Private Sub cbonmDocHistEsc_DropDown()
    Dim RsTMP As Recordset
    cbonmDocHistEsc.Clear
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Documentos ORDER BY Descr")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                cbonmDocHistEsc.AddItem RsTMP.Fields("Descr")
                RsTMP.MoveNext
            Loop
            RsTMP.Close
    End If
End Sub

Private Sub chkUsarHistEscImp_Click()
    If chkUsarHistEscImp.Value = 0 Then
            cbonmDocHistEsc.Clear
            cbonmDocHistEsc.Enabled = False
        Else
            cbonmDocHistEsc.Enabled = True
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
    SST_Config.Tab = 0
    Set RsUnid = BD.OpenRecordset("SELECT * FROM Unidades")
    If RsUnid.BOF And RsUnid.EOF Then
            MsgBox "Cadastre primeiro a Unidade de Ensino", vbInformation, "CESNet - Aviso!"
        Else
            CB_Unidade.Clear
            RsUnid.MoveFirst
            Do Until RsUnid.EOF
                CB_Unidade.AddItem (RsUnid.Fields("UnidID"))
                RsUnid.MoveNext
            Loop
    End If
    Set RsConfig = BD.OpenRecordset("SELECT * FROM Config")
    RsConfig.MoveFirst
    Meb_NotaMedia.Text = IIf(IsNull(RsConfig.Fields("NotaMedia")), "60", RsConfig.Fields("NotaMedia"))
    Meb_MaxAcessos.Text = IIf(IsNull(RsConfig.Fields("MaxAcessos")), "100", RsConfig.Fields("MaxAcessos"))
    Meb_MaxModulos.Text = IIf(IsNull(RsConfig.Fields("MaxModulosEmpr")), "1", RsConfig.Fields("MaxModulosEmpr"))
    
    chkCartConjugada.Value = IIf(IsNull(RsConfig.Fields("CartConjugada")), "0", RsConfig.Fields("CartConjugada"))
    
    Cb_VincMod.Text = IIf(RsConfig.Fields("VincModulos") = True, "SIM", "NAO")
    Opt_SisNota(IIf(RsConfig.Fields("SisNota") = True, 0, 1)).Value = True
    Txt_NumRepro.Text = IIf(IsNull(RsConfig.Fields("NumRepro")), "0", RsConfig.Fields("NumRepro"))
    
    Txt_InatDias.Text = IIf(IsNull(RsConfig.Fields("InatDias")), "5", RsConfig.Fields("InatDias"))
    Chk_InatBloc.Value = IIf(RsConfig.Fields("BloquearInativo") = True, 1, 0)
    'Txt_LocalBD.Text = PathBD
    Txt_RPTime.Text = IIf(IsNull(RsConfig.Fields("RPTime")), 1, RsConfig.Fields("RPTime"))
    Txt_RPMax.Text = IIf(IsNull(RsConfig.Fields("RPMax")), 1, RsConfig.Fields("RPMax"))
    Txt_RPMaxReg.Text = IIf(IsNull(RsConfig.Fields("RPMaxReg")), 1, RsConfig.Fields("RPMaxReg"))
    
    chkUsarHistEscImp.Value = IIf(IsNull(RsConfig.Fields("UsarHistEscImp")), 0, RsConfig.Fields("UsarHistEscImp"))
    chkBloqRenovVencida.Value = IIf(IsNull(RsConfig.Fields("BloqRenovVencida")), 0, RsConfig.Fields("BloqRenovVencida"))
    
    If chkUsarHistEscImp.Value = 0 Then
            cbonmDocHistEsc.Enabled = False
        Else
            cbonmDocHistEsc.Enabled = True
    End If
    If Not IsNull(RsConfig.Fields("nmDocHistEsc")) Then
        cbonmDocHistEsc.Clear
        cbonmDocHistEsc.AddItem IIf(PgDescArqDoc(RsConfig.Fields("nmDocHistEsc")) = "", " ", PgDescArqDoc(RsConfig.Fields("nmDocHistEsc")))
        cbonmDocHistEsc.Text = cbonmDocHistEsc.List(0)
    End If
    
    'Txt_pgNumMinDiscipl.Text = IIf(IsNull(RsConfig.Fields("pgNumMinDiscipl")), "0", RsConfig.Fields("pgNumMinDiscipl"))
    Chk_DeslWin.Value = IIf(RsConfig.Fields("DeslWin") = True, 1, 0)
    CB_Unidade.Text = Mid(String(3, "0"), 1, 3 - Len(RsConfig.Fields("Unidade"))) & RsConfig.Fields("Unidade")
    RsUnid.FindFirst "UnidID = '" & CB_Unidade.Text & "'"
    If RsUnid.NoMatch Then
            MsgBox "Erro na procura pela Unidade de Ensino. Por favor Verifique", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            Lb_Unid.Caption = RsUnid.Fields("Nome")
    End If
    
End Sub

Private Sub Meb_MaxAcessos_GotFocus()
    Meb_MaxAcessos.SelStart = 0
    Meb_MaxAcessos.SelLength = 5
End Sub

Private Sub Meb_NotaMedia_GotFocus()
    Meb_NotaMedia.SelStart = 0
    Meb_NotaMedia.SelLength = 4
End Sub





Private Sub Txt_InatDias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii <= 47 Or KeyAscii >= 58 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_NumMinDiscipl_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii <= 47 Or KeyAscii >= 58 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_NumRepro_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii <= 47 Or KeyAscii >= 58 Then
        KeyAscii = 0
    End If
End Sub



'Private Sub Txt_RPTime_Change()
'    If Val(Txt_RPTime.Text) > 59 Then
'        MsgBox "O intervalo deve ser entre 0 e 59", vbInformation, "CESNet - Aviso!"
'        Txt_RPTime.Text = ""
'    End If
'End Sub

Private Sub Txt_RPTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii <= 47 Or KeyAscii >= 58 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_RPMax_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii <= 47 Or KeyAscii >= 58 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Txt_RPMaxReg_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii <= 47 Or KeyAscii >= 58 Then
        KeyAscii = 0
    End If
End Sub

