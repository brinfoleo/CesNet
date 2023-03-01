VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_PreMatr 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Pre-Matricula"
   ClientHeight    =   7035
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9060
   Icon            =   "Form_PreMatr.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7035
   ScaleWidth      =   9060
   Begin VB.CommandButton Bt_Gravar 
      Caption         =   "&Gravar Matricula"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7245
      TabIndex        =   12
      Top             =   1200
      Width           =   1770
   End
   Begin VB.CommandButton Bt_Incluir 
      Caption         =   "Incluir Matricula"
      Height          =   555
      Left            =   7245
      TabIndex        =   9
      Top             =   495
      Width           =   1770
   End
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "Cancelar"
      Enabled         =   0   'False
      Height          =   555
      Left            =   7245
      TabIndex        =   16
      Top             =   1740
      Width           =   1770
   End
   Begin VB.Frame Frame1 
      Height          =   1845
      Left            =   45
      TabIndex        =   2
      Top             =   405
      Width           =   7080
      Begin VB.TextBox txtNome 
         Height          =   285
         Left            =   840
         MaxLength       =   50
         TabIndex        =   22
         Top             =   600
         Width           =   5715
      End
      Begin VB.TextBox txtUnidade 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         TabIndex        =   10
         Top             =   990
         Width           =   4335
      End
      Begin VB.TextBox txtMatrOld 
         Enabled         =   0   'False
         Height          =   285
         Left            =   840
         MaxLength       =   15
         TabIndex        =   11
         Top             =   1380
         Width           =   1875
      End
      Begin MSComCtl2.DTPicker DTP_Dt 
         Height          =   285
         Left            =   840
         TabIndex        =   6
         Top             =   225
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         Format          =   16580609
         CurrentDate     =   38376
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Unidade:"
         Height          =   195
         Left            =   15
         TabIndex        =   15
         Top             =   1035
         Width           =   735
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Nome:"
         Height          =   240
         Left            =   195
         TabIndex        =   14
         Top             =   615
         Width           =   555
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Matricula Antiga:"
         Height          =   420
         Left            =   60
         TabIndex        =   7
         Top             =   1290
         Width           =   690
      End
      Begin VB.Label Label7 
         Caption         =   "Data:"
         Height          =   240
         Left            =   285
         TabIndex        =   5
         Top             =   270
         Width           =   465
      End
      Begin VB.Label lbMatricula 
         Alignment       =   2  'Center
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "00.000.0000"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5220
         TabIndex        =   4
         Top             =   180
         Width           =   1725
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Matricula:"
         Height          =   195
         Left            =   4500
         TabIndex        =   3
         Top             =   270
         Width           =   690
      End
   End
   Begin VB.Frame Frame2 
      Height          =   4470
      Left            =   60
      TabIndex        =   0
      Top             =   2460
      Width           =   8925
      Begin VB.Frame Frame3 
         Height          =   1275
         Left            =   120
         TabIndex        =   18
         Top             =   3060
         Width           =   8595
         Begin VB.ComboBox cbProva 
            Height          =   315
            Left            =   1500
            Style           =   2  'Dropdown List
            TabIndex        =   21
            Top             =   180
            Width           =   6975
         End
         Begin VB.CommandButton btAplicar 
            Caption         =   "&Aplicar"
            Height          =   435
            Left            =   6420
            TabIndex        =   20
            Top             =   720
            Width           =   2055
         End
         Begin VB.Label Label4 
            Alignment       =   1  'Right Justify
            Caption         =   "Num. da Prova:"
            Height          =   255
            Left            =   60
            TabIndex        =   19
            Top             =   300
            Width           =   1215
         End
      End
      Begin MSFlexGridLib.MSFlexGrid msfgDiscipl 
         Height          =   2655
         Left            =   120
         TabIndex        =   17
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   4683
         _Version        =   393216
         FixedCols       =   0
         SelectionMode   =   1
         FormatString    =   $"Form_PreMatr.frx":030A
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
      Begin VB.ComboBox cb_Ensino 
         Height          =   315
         Left            =   900
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   0
         Width           =   2895
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "Ensino:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   135
         TabIndex        =   1
         Top             =   60
         Width           =   780
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "PRE-MATRICULA"
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
      Top             =   0
      Width           =   9105
   End
End
Attribute VB_Name = "Form_PreMatr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim EnsinoID            As Integer
Dim DisciplinaID        As Integer
'Abrir conexoes externas

Dim Fora                As Boolean 'True-pegando BD / false - Sem bd
Dim ForaDP(30)          As String
Dim ForaLocBD           As String




Private Sub LimpForm()
    lbMatricula.Caption = "00.000.0000"
    txtNome.Text = ""
    txtMatrOld.Text = ""
    Cb_Ensino.Clear
    msfgDiscipl.Rows = 1
    msfgDiscipl.Rows = 2
    cbProva.Clear
End Sub
Private Sub GRVDisciplina(MatrID As String)
    Dim RsTMP           As Recordset
    'dim EnsinoID        As Integer
    Dim n               As Integer
    Dim nProva          As String
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina") 'WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
    
    For n = 1 To msfgDiscipl.Rows - 1
        
        nProva = left(msfgDiscipl.TextMatrix(n, 1), 3)
        
        If Trim(nProva) <> "" Then
            DisciplinaID = PgIDDisciplina(msfgDiscipl.TextMatrix(n, 0))
            RsTMP.AddNew
            RsTMP.Fields("MatrID") = MatrID
            RsTMP.Fields("EnsinoID") = EnsinoID
            RsTMP.Fields("DisciplinaID") = DisciplinaID
            RsTMP.Fields("UsuarioID") = UsuarioID
            RsTMP.Fields("DtHrSis") = Now()
            RsTMP.Update
            '========================================================
            Call GRVSerie(MatrID, DisciplinaID, nProva)
            '========================================================
        End If
        
    Next
    RsTMP.Close
End Sub


Private Sub GrvEnsino(MatrID As String)
    Dim RsTMP       As Recordset
    'dim EnsinoID    As Integer
    
    'EnsinoID = PgIDEnsino(cb_Ensino.Text)
    Set RsTMP = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsTMP.BOF And RsTMP.EOF Then
        RsTMP.AddNew
        RsTMP.Fields("MatrID") = MatrID
        RsTMP.Fields("EnsinoID") = EnsinoID
        RsTMP.Fields("DtInicio") = DTP_Dt.Value
        RsTMP.Fields("UsuarioID") = UsuarioID
        RsTMP.Fields("DtHrSis") = Now()
        RsTMP.Update
    End If
    RsTMP.Close
End Sub


Private Sub GRVSerie(MatrID As String, DisciplinaID As Integer, nProva As String)
    Dim RsGradeSerie    As Recordset
    Dim RsProva         As Recordset
    Dim RsMatrSerie     As Recordset
    Dim RsMatrProva     As Recordset
    
    'dim EnsinoID        As Integer
    Dim SerieID         As Integer
    
    
    If nProva = "" Then Exit Sub
    
    
    'PEGA A SERIE DA PROVA =================================================================
    Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE EnsinoID = " & EnsinoID & _
                                        " AND DisciplinaID = " & DisciplinaID & _
                                        " AND NProva = '" & nProva & "'")
    If RsProva.BOF And RsProva.EOF Then
            MsgBox "Erro ao localizar GRADE_ENSINO_SERIE (GRVSerie)", vbInformation, "CESNet - Aviso"
            RsProva.Close
            Exit Sub
        Else
            RsProva.MoveFirst
            
            SerieID = RsProva.Fields("SerieID")
            
            RsProva.Close
    End If
    '======================================================================================
    
    'PEGAR GRADE DE SERIES ================================================================
    Set RsGradeSerie = BD.OpenRecordset("SELECT * FROM GradeEnsinoSeries WHERE" & _
                                        " EnsinoID = " & EnsinoID & _
                                        " AND DisciplinaID = " & DisciplinaID & _
                                        " ORDER BY SerieID")
    
    
    RsGradeSerie.FindFirst "SerieID = " & SerieID 'BUSCAR A SERIE
    
    'ABRIR A TAB. MATRICULASERIE ===========================================================
    Set RsMatrSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie")
    '=======================================================================================
    
    'GRAVA AS SERIES===========================================================
    Do Until RsGradeSerie.EOF
        RsMatrSerie.AddNew
        RsMatrSerie.Fields("MatrID") = MatrID
        RsMatrSerie.Fields("EnsinoID") = EnsinoID
        RsMatrSerie.Fields("DisciplinaID") = DisciplinaID
        RsMatrSerie.Fields("SerieID") = RsGradeSerie.Fields("SerieID")
        'RsMatrSerie.Fields("DtIni") = Date
        RsMatrSerie.Fields("UsuarioID") = UsuarioID
        RsMatrSerie.Fields("DtHrSis") = Now()
        RsMatrSerie.Update
        RsGradeSerie.MoveNext
    Loop
    '===========================================================================
    
    'PEGA A SERIE DA PROVA =================================================================
    Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE EnsinoID = " & EnsinoID & _
                                        " AND DisciplinaID = " & DisciplinaID & _
                                        " AND SerieID = " & SerieID & _
                                        " ORDER BY NProva")
    RsProva.MoveFirst
                                        
    '=======================================================================================
    Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva")
    
    'GARVAR AS PROVAS ===================================================================
    Do Until RsProva.Fields("NProva") = nProva
        RsMatrProva.AddNew
        RsMatrProva.Fields("MatrID") = MatrID
        RsMatrProva.Fields("EnsinoID") = EnsinoID
        RsMatrProva.Fields("DisciplinaID") = DisciplinaID
        RsMatrProva.Fields("NProva") = RsProva.Fields("NProva")
        RsMatrProva.Fields("Assunto") = RsProva.Fields("Assunto")
        RsMatrProva.Fields("Nota") = "100"
        RsMatrProva.Fields("Tipo") = "A"
        RsMatrProva.Fields("DtAvaliacao") = DTP_Dt.Value
        RsMatrProva.Fields("Aprovado") = True
        RsMatrProva.Fields("Status") = "HB"
        RsMatrProva.Fields("DtHrAv") = Now
        RsMatrProva.Fields("ProfIDN") = "0"
        RsMatrProva.Fields("DtHrN") = Now
        RsMatrProva.Fields("UsuarioIDN") = UsuarioID
        RsMatrProva.Update
        
        RsProva.MoveNext
    Loop
    '======================================================================================
                                        
   
End Sub


Private Sub LstDisciplinas()
    Dim RsTMP       As Recordset
    'dim EnsinoID    As Integer
    
    EnsinoID = PgIDEnsino(Cb_Ensino.Text)
    msfgDiscipl.Rows = 1
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID)
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                msfgDiscipl.Rows = msfgDiscipl.Rows + 1
                msfgDiscipl.TextMatrix(msfgDiscipl.Rows - 1, 0) = PgNomeDisciplina(RsTMP.Fields("DisciplinaID"))
                If Fora = True Then
                    If Fora_PgDisciplinaConcluida(RsTMP.Fields("DisciplinaID")) = True Then
                        'Disciplina concluida
                         msfgDiscipl.TextMatrix(msfgDiscipl.Rows - 1, 1) = "CONCLUIDA"
                    End If
                End If
                RsTMP.MoveNext
            Loop
        
    End If
End Sub


Private Sub btAplicar_Click()
    With msfgDiscipl
        .TextMatrix(.Row, 1) = cbProva.Text
        cbProva.Clear
    End With
End Sub



Private Sub Bt_Cancelar_Click()
 
    HDForm (False)
    LimpForm
End Sub
Private Sub HDForm(op As Boolean)
    DTP_Dt.Enabled = op
    txtMatrOld.Enabled = op
    txtNome.Enabled = op
    txtMatrOld.Enabled = op
    Cb_Ensino.Enabled = op
    msfgDiscipl.Enabled = op
    cbProva.Enabled = op
    btAplicar.Enabled = op
    
    Bt_Incluir.Enabled = IIf(op = True, False, True)
    
    Bt_Gravar.Enabled = op
    Bt_Cancelar.Enabled = op

End Sub
Private Sub Bt_Gravar_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    If ChkAcesso(Me.Name, "A") = False Then Exit Sub
    Dim MatrID As String
    
    If Trim(txtNome.Text) = "" Then
        MsgBox "O campo NOME nao pode ser deixado em branco!!!", vbInformation, "Atenção"
        Exit Sub
    End If
    If ValidarSoftware("Matriculas") = False Then Exit Sub
    MatrID = GRVMatricula
    If Trim(MatrID) = "" Then
        MsgBox "Erro ao gravar Dados Pessoais"
        Exit Sub
    End If
    
    Call GrvEnsino(MatrID)
    Call GRVDisciplina(MatrID)
    
    HDForm (False)
End Sub

Private Function GRVMatricula() As String
    Dim strMatrID As String
    strMatrID = Right(DTP_Dt.Value, 2) & "." & left(txtUnidade.Text, 3)
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID LIKE '" & strMatrID & "*'")
    With RsTMP
        If .BOF And .EOF Then
                strMatrID = strMatrID & ".0001"
            Else
                .MoveLast
                strMatrID = strMatrID & "." & Mid(String(4, "0"), 1, 4 - Len(Right(.Fields("MatrID") + 1, 4))) & Right(.Fields("MatrID") + 1, 4)

        End If
    End With

    
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas")
    RsMatricula.AddNew
    RsMatricula.Fields("MatrID") = strMatrID
    RsMatricula.Fields("DtMat") = DTP_Dt.Value
    RsMatricula.Fields("Nome") = Trim(UCase(txtNome.Text))
    RsMatricula.Fields("Unidade") = UnidadeEnsino & " - " & UnidadeEnsinoNome
    RsMatricula.Fields("UnidadeID") = UnidadeEnsino
    
    If Fora = True Then
        RsMatricula.Fields("End") = IIf(Trim(ForaDP(1)) = "", Null, ForaDP(1))
        RsMatricula.Fields("Bai") = IIf(Trim(ForaDP(2)) = "", Null, ForaDP(2))
        RsMatricula.Fields("Mun") = IIf(Trim(ForaDP(3)) = "", Null, ForaDP(3))
        RsMatricula.Fields("UF") = IIf(Trim(ForaDP(4)) = "", Null, ForaDP(4))
        RsMatricula.Fields("CEP") = IIf(Trim(ForaDP(5)) = "", Null, ForaDP(5))
        RsMatricula.Fields("Tel1") = IIf(Trim(ForaDP(6)) = "", Null, ForaDP(6))
        RsMatricula.Fields("RG") = IIf(Trim(ForaDP(7)) = "", Null, ForaDP(7))
        RsMatricula.Fields("OE") = IIf(Trim(ForaDP(9)) = "", Null, ForaDP(9))
        RsMatricula.Fields("CPF") = IIf(Trim(ForaDP(10)) = "", Null, ForaDP(10))
        RsMatricula.Fields("Sexo") = IIf(Trim(ForaDP(11)) = "", Null, ForaDP(11))
        RsMatricula.Fields("Nasc") = IIf(Trim(ForaDP(12)) = "", Null, ForaDP(12))
        RsMatricula.Fields("EstCivil") = IIf(Trim(ForaDP(13)) = "", Null, ForaDP(13))
        RsMatricula.Fields("Nacion") = IIf(Trim(ForaDP(14)) = "", Null, ForaDP(14))
        RsMatricula.Fields("Natural") = IIf(Trim(ForaDP(15)) = "", Null, ForaDP(15))
        RsMatricula.Fields("NaturalUF") = IIf(Trim(ForaDP(16)) = "", Null, ForaDP(16))
        RsMatricula.Fields("Raca") = IIf(Trim(ForaDP(17)) = "", Null, ForaDP(17))
        RsMatricula.Fields("Pai") = IIf(Trim(ForaDP(18)) = "", Null, ForaDP(18))
        RsMatricula.Fields("Mae") = IIf(Trim(ForaDP(19)) = "", Null, ForaDP(19))
        RsMatricula.Fields("Obs") = IIf(Trim(ForaDP(20)) = "", Null, ForaDP(20))
        '***************************************************************************************************
            'ForaDP(8) = IIf(IsNull(Rst.Fields("RG_emissao")), "", Rst.Fields("RG_emissao"))
            'ForaDP(21) = IIf(IsNull(Rst.Fields("Data_Cadastro")), "", Rst.Fields("Data_Cadastro"))
            'ForaDP(29) = IIf(IsNull(Rst.Fields("Grau")), "", Rst.Fields("Grau"))
            'ForaDP(30) = Rst.Fields("Matricula")
        '/****************************************************************************************************
        txtMatrOld.Text = UCase(txtMatrOld.Text)
    End If
    
    
    
    RsMatricula.Fields("NumAnt") = IIf(Trim(UCase(txtMatrOld.Text)) = "", Null, Trim(txtMatrOld.Text))
    RsMatricula.Fields("UsuarioID") = UsuarioID
    RsMatricula.Fields("DtHrSis") = Now()
    RsMatricula.Update
    lbMatricula.Caption = strMatrID
    GRVMatricula = strMatrID
End Function

Private Sub Bt_Incluir_Click()
    If ChkAcesso(Me.Name, "N") = False Then Exit Sub
    LimpForm
    HDForm (True)
End Sub




Private Sub Cb_Ensino_Click()
    EnsinoID = PgIDEnsino(Cb_Ensino.Text)
    LstDisciplinas
End Sub

Private Sub Cb_Ensino_DropDown()
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









Private Sub cbProva_DropDown()
    cbProva.Clear
    If EnsinoID = 0 Or DisciplinaID = 0 Then
        MsgBox "Selecione um Ensino ou Disciplina", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY NProva")
    If RsProvas.BOF And RsProvas.EOF Then
            MsgBox "Nenhuma prova encontrada"
            Exit Sub
        Else
            RsProvas.MoveFirst
            
            Do Until RsProvas.EOF
                cbProva.AddItem (RsProvas.Fields("NProva") & " - " & RsProvas.Fields("Assunto"))
                RsProvas.MoveNext
            Loop
    End If
End Sub






Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "N") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Fora = False
    DTP_Dt.Value = Date
    txtUnidade.Text = UnidadeEnsino & " - " & UnidadeEnsinoNome
    HDForm (False)
    Fora_chkArquivo
End Sub
Private Sub Fora_chkArquivo()
    If Dir(App.path & "\CESNet.ext") <> "" Then
            'Fora = True
            'Fora_PgBDImport
            'If ForaLocBD = "" Then
            '        MsgBox "Erro ao localizar banco de dados auxiliar.", vbInformation, "Aviso"
            '        Fora = False
            '        Exit Sub
            '    Else
                   Fora = ImportOpenBD
            'End If
        Else
            Fora = False
    End If
End Sub

Private Sub msfgDiscipl_Click()
    DisciplinaID = PgIDDisciplina(msfgDiscipl.TextMatrix(msfgDiscipl.Row, 0))
End Sub

Private Sub txtMatrOld_Change()
    If Fora = True Then
        Fora_PgDados
    End If
End Sub

Private Sub txtMatrOld_KeyDown(KeyCode As Integer, Shift As Integer)


    If KeyCode = 114 Then
       formBuscar.IniciarBusca "ALUNOS", , , , "fora"
    End If
End Sub

Private Sub txtNome_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub

Private Function Fora_PgDisciplinaConcluida(DisciplID As Integer) As Boolean
    'True = Disciplina Concluida
    'False = Disciplina Cursando
    Dim sSQL    As String
    Dim Rst     As New ADODB.Recordset
    Dim Disc    As String
    Dim IDProva As String
    
    Disc = PgNomeDisciplina(DisciplID)
    
    'Pega o ID da ultima prova
    sSQL = "SELECT * FROM Modulos WHERE GRAU ='" & ForaDP(29) & "' AND Nome = '" & Disc & "' ORDER BY Modulo"
    Rst.Open sSQL, conexao, adOpenStatic ' adOpenDynamic
    If Rst.BOF And Rst.EOF Then
            Fora_PgDisciplinaConcluida = False
            Rst.Close
            Exit Function
        Else
            Rst.MoveLast
            IDProva = Rst.Fields("Codigo")
    End If
    Rst.Close
    
    'Pega as provas efetuadas
    sSQL = "SELECT * FROM Avaliacoes WHERE Codigo_Aluno=" & ForaDP(30) & " AND CODIGO_Modulo = " & IDProva & " ORDER BY Codigo"
    Rst.Open sSQL, conexao, adOpenStatic
    If Rst.BOF And Rst.EOF Then
            Fora_PgDisciplinaConcluida = False
            Rst.Close
            Exit Function
        Else
            Rst.MoveLast
            IDProva = Rst.Fields("Codigo")
    End If
    Rst.Close
    
End Function

Private Sub Fora_PgDados()
    On Error Resume Next
    Dim Rst     As New ADODB.Recordset
    Dim sSQL    As String
    sSQL = "SELECT * FROM ALUNOS WHERE MATRICULA_ANTIGA='" & UCase(txtMatrOld.Text) & "'"
    
     Rst.Open sSQL, conexao
    
    If Rst.BOF And Rst.EOF Then
            Rst.Close
            txtNome.Text = ""
            Cb_Ensino.Clear
            msfgDiscipl.Rows = 1
            Exit Sub
        Else
            Rst.MoveFirst
            ForaDP(0) = IIf(IsNull(Rst.Fields("Nome")), "", Rst.Fields("Nome"))
            ForaDP(1) = IIf(IsNull(Rst.Fields("Endereco")), "", Rst.Fields("Endereco"))
            ForaDP(2) = IIf(IsNull(Rst.Fields("Bairro")), "", Rst.Fields("Bairro"))
            ForaDP(3) = IIf(IsNull(Rst.Fields("Municipio")), "", Rst.Fields("Municipio"))
            ForaDP(4) = IIf(IsNull(Rst.Fields("Estado")), "", Rst.Fields("Estado"))
            ForaDP(5) = IIf(IsNull(Rst.Fields("CEP")), "", Rst.Fields("CEP"))
            ForaDP(6) = IIf(IsNull(Rst.Fields("Telefone")), "", Rst.Fields("Telefone"))
            ForaDP(7) = IIf(IsNull(Rst.Fields("RG")), "", Rst.Fields("RG"))
            ForaDP(8) = IIf(IsNull(Rst.Fields("RG_emissao")), "", Rst.Fields("RG_emissao"))
            ForaDP(9) = IIf(IsNull(Rst.Fields("RG_Orgao_Exp")), "", Rst.Fields("RG_Orgao_Exp"))
            ForaDP(10) = IIf(IsNull(Rst.Fields("CPF")), "", Rst.Fields("CPF"))
            ForaDP(11) = IIf(IsNull(Rst.Fields("Sexo")), "", Rst.Fields("Sexo"))
            ForaDP(12) = IIf(IsNull(Rst.Fields("Nascimento")), "", Rst.Fields("Nascimento"))
            ForaDP(13) = IIf(IsNull(Rst.Fields("Estado_Civil")), "", Rst.Fields("Estado_Civil"))
            ForaDP(14) = IIf(IsNull(Rst.Fields("Nacionalidade")), "", Rst.Fields("Nacionalidade"))
            ForaDP(15) = IIf(IsNull(Rst.Fields("Naturalidade_Municipio")), "", Rst.Fields("Naturalidade_Municipio"))
            ForaDP(16) = IIf(IsNull(Rst.Fields("Naturalidade_Estado")), "", Rst.Fields("Naturalidade_Estado"))
            ForaDP(17) = IIf(IsNull(Rst.Fields("COR")), "", Rst.Fields("COR"))
            ForaDP(18) = IIf(IsNull(Rst.Fields("Pai")), "", Rst.Fields("Pai"))
            ForaDP(19) = IIf(IsNull(Rst.Fields("Mae")), "", Rst.Fields("Mae"))
            ForaDP(20) = IIf(IsNull(Rst.Fields("Observacoes")), "", Rst.Fields("Observacoes"))
            ForaDP(21) = IIf(IsNull(Rst.Fields("Data_Cadastro")), "", Rst.Fields("Data_Cadastro"))
            ForaDP(29) = IIf(IsNull(Rst.Fields("Grau")), "", Rst.Fields("Grau"))
            ForaDP(30) = Rst.Fields("Matricula")
            Rst.Close
            
            txtNome.Text = ForaDP(0)
            If Trim(ForaDP(29)) <> "" Then
                Cb_Ensino.Clear
                Cb_Ensino.AddItem PgNomeEnsino(CInt(ForaDP(29)))
                Cb_Ensino.Text = Cb_Ensino.List(0)
                Cb_Ensino_Click
            End If
    End If

End Sub


