VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_ResultProvas 
   BorderStyle     =   0  'None
   Caption         =   "CESNet - Resultado das Provas"
   ClientHeight    =   4395
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12570
   Icon            =   "Form_ResultProvas.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   4395
   ScaleWidth      =   12570
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame_Leg 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   8940
      TabIndex        =   7
      Top             =   3240
      Width           =   2955
      Begin VB.Label Label5 
         Caption         =   "NC = Não Corrigido"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   10
         Top             =   780
         Width           =   2475
      End
      Begin VB.Label Label4 
         Caption         =   "NH = Não Habilitado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   480
         Width           =   2535
      End
      Begin VB.Label Label3 
         Caption         =   "HB = Habilitado"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   8
         Top             =   180
         Width           =   2235
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFG_Result 
      Height          =   1860
      Left            =   135
      TabIndex        =   1
      Top             =   1305
      Width           =   10425
      _ExtentX        =   18389
      _ExtentY        =   3281
      _Version        =   393216
      Cols            =   5
      FixedCols       =   0
      GridLines       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"Form_ResultProvas.frx":030A
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSMask.MaskEdBox MebMatricula 
      Height          =   375
      Left            =   1440
      TabIndex        =   3
      Top             =   600
      Width           =   1995
      _ExtentX        =   3519
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##.###.####"
      PromptChar      =   "_"
   End
   Begin VB.Label Lb_Msg2 
      Caption         =   "Pressione / (barra) para limpar a tela."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   495
      Left            =   180
      TabIndex        =   6
      Top             =   3780
      Width           =   8715
   End
   Begin VB.Label Lb_Msg1 
      Caption         =   "Digite o seu número de matricula."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   180
      TabIndex        =   5
      Top             =   3240
      Width           =   8715
   End
   Begin VB.Label Lb_Nome 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3540
      TabIndex        =   4
      Top             =   600
      Width           =   7215
   End
   Begin VB.Label Label2 
      Caption         =   "Matricula:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Top             =   600
      Width           =   1275
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00FF0000&
      Caption         =   "RESULTADO DAS PROVAS"
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
      Width           =   7215
   End
End
Attribute VB_Name = "Form_ResultProvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim MatrID As String
Dim RsMatrProva As Recordset



Private Sub Form_Activate()
    MebMatricula.SetFocus
End Sub

Private Sub MebMatricula_Change()
      If IsNumeric(MebMatricula) Then
        Lb_Nome.Caption = PgDadosMatr(MebMatricula.Text).Nome
        If Lb_Nome.Caption = "" Then
                MSFG_Result.Rows = 1
                Exit Sub
            Else
                MatrID = MebMatricula.Text
                MstProvas
        End If
    End If

End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub

Private Sub Form_Resize()
    Label1.Width = Form_ResultProvas.ScaleWidth
    'Lb_Disciplina.Width = Label1.Width
    MSFG_Result.left = 200
    MSFG_Result.Width = Form_ResultProvas.ScaleWidth - 200
    MSFG_Result.Height = Form_ResultProvas.ScaleHeight - (Lb_Msg1.Height + Lb_Msg2.Height + 1800)
    'Bt_Fechar.left = Form_ResultProvas.ScaleWidth - 300
    
    Lb_Msg1.top = Form_ResultProvas.ScaleHeight - (Lb_Msg1.Height + Lb_Msg2.Height)
    Lb_Msg2.top = Form_ResultProvas.ScaleHeight - Lb_Msg2.Height
    
    Frame_Leg.top = Form_ResultProvas.ScaleHeight - (Frame_Leg.Height + 100)
End Sub

Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    If KeyAscii = 16 Then 'CTRL + P
        'Autentica o usuario
        'If Form_AutenticacaoUsuario.CarregarForm = True Then
            Unload Me
            Exit Sub
        'End If
    End If
    If KeyAscii = 13 Then
        Lb_Nome.Caption = PgDadosMatr(MebMatricula.Text).Nome
        If Lb_Nome.Caption = "" Then
                MSFG_Result.Rows = 1
                Exit Sub
            Else
                MatrID = MebMatricula.Text
                MstProvas
        End If
    End If
    If KeyAscii = 27 Or KeyAscii = 47 Then
        MebMatricula.PromptInclude = False
        MebMatricula.Text = ""
        MebMatricula.PromptInclude = True
        Lb_Nome.Caption = ""
        MSFG_Result.Rows = 1
    End If
        
End Sub
Private Sub MstProvas()
    DoEvents
    Dim RsTMP As Recordset
    
    Dim IDProva As Variant
    
    MSFG_Result.Rows = 1
    Set RsMatrProva = BD.OpenRecordset("SELECT TOP " & RPMax & " * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' ORDER BY DtAvaliacao DESC")
    If RsMatrProva.BOF And RsMatrProva.EOF Then
        
        Else
            RsMatrProva.MoveLast
            Do Until RsMatrProva.BOF
                With MSFG_Result
                    If IsNull(RsMatrProva.Fields("DisciplinaID")) Then
                        Else
                            .Rows = .Rows + 1
                            IDProva = RsMatrProva.Fields("ID")
                            .TextMatrix(.Rows - 1, 0) = RsMatrProva.Fields("DtAvaliacao")
                            .TextMatrix(.Rows - 1, 1) = PgNomeDisciplina(RsMatrProva.Fields("DisciplinaID"))
                            .TextMatrix(.Rows - 1, 2) = RsMatrProva.Fields("NProva")
                            
                            'If Trim(RsMatrProva.Fields("Nota")) <> "" Then
                                    '.TextMatrix(.Rows - 1, 3) = IIf(RsMatrProva.Fields("Aprovado") = True, "HB", "NH")
                                'Else
                            If SisNota = True Then
                                    .TextMatrix(.Rows - 1, 3) = RsMatrProva.Fields("Nota")
                                Else
                                    .TextMatrix(.Rows - 1, 3) = RsMatrProva.Fields("Status")
                            End If
                            Set RsTMP = BD.OpenRecordset("SELECT * FROM ProvasTMP WHERE MPID = " & IDProva & " AND Obs <> Null ORDER BY Seq")
                            If RsTMP.BOF And RsTMP.EOF Then
                                Else
                                    RsTMP.MoveLast
                                    .TextMatrix(.Rows - 1, 4) = RsTMP.Fields("Obs")
                            End If
                            'End If
                    End If
                    RsMatrProva.MovePrevious
                End With
            Loop
    End If
    MSFG_Result.Row = 0
    RsMatrProva.Close
    MebMatricula.SetFocus
    MebMatricula.SelStart = 0
        
End Sub

Private Sub MebMatricula_LostFocus()
    MebMatricula.SetFocus
End Sub
