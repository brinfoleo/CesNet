VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form_CoordImpCert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Configurar Impressão de Coordenadas"
   ClientHeight    =   5070
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8520
   Icon            =   "Form_CoodImpCert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5070
   ScaleWidth      =   8520
   Begin VB.TextBox txtFormEstudo 
      Height          =   315
      Left            =   6540
      MaxLength       =   20
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   1620
      Width           =   1815
   End
   Begin VB.Frame Frame1 
      Height          =   1575
      Left            =   5040
      TabIndex        =   9
      Top             =   2460
      Width           =   3375
      Begin VB.TextBox txtNH 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   14
         Top             =   1140
         Width           =   1935
      End
      Begin VB.TextBox txtHB 
         Height          =   285
         Left            =   1320
         MaxLength       =   60
         TabIndex        =   13
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "Nomeclatura para informar a situação final da disciplina."
         Height          =   375
         Left            =   120
         TabIndex        =   12
         Top             =   180
         Width           =   2835
      End
      Begin VB.Label Label3 
         Caption         =   "Não Habilitado:"
         Height          =   195
         Left            =   120
         TabIndex        =   11
         Top             =   1140
         Width           =   1095
      End
      Begin VB.Label Label2 
         Caption         =   "Habilitados:"
         Height          =   195
         Left            =   480
         TabIndex        =   10
         Top             =   780
         Width           =   735
      End
   End
   Begin VB.CheckBox chkUsarSigla 
      Caption         =   "Usar sigla da Inst. Ensino"
      Height          =   315
      Left            =   5040
      TabIndex        =   8
      Top             =   780
      Width           =   3075
   End
   Begin VB.CheckBox chkUsarCidRed 
      Caption         =   "Usar descrição de cidade reduzida."
      Height          =   315
      Left            =   5040
      TabIndex        =   7
      Top             =   1140
      Width           =   3135
   End
   Begin VB.TextBox Txt_Coord 
      Appearance      =   0  'Flat
      BackColor       =   &H0080FFFF&
      BorderStyle     =   0  'None
      Height          =   225
      Left            =   2460
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2040
      Visible         =   0   'False
      Width           =   915
   End
   Begin VB.ComboBox Cb_Modelo 
      Height          =   315
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   420
      Width           =   3915
   End
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   735
      Left            =   6780
      Picture         =   "Form_CoodImpCert.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4200
      Width           =   1575
   End
   Begin VB.CommandButton Bt_Gravar 
      Caption         =   "&Gravar"
      Height          =   735
      Left            =   5100
      Picture         =   "Form_CoodImpCert.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4200
      Width           =   1575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFG_CoordCert 
      Height          =   3180
      Left            =   0
      TabIndex        =   1
      Top             =   840
      Width           =   4920
      _ExtentX        =   8678
      _ExtentY        =   5609
      _Version        =   393216
      Cols            =   3
      AllowUserResizing=   1
      FormatString    =   "^CAMPO                                  |^TOPO            |^MARG. ESQ. "
   End
   Begin VB.Label Label5 
      Alignment       =   1  'Right Justify
      Caption         =   "Forma de Estudo:"
      Height          =   195
      Left            =   5040
      TabIndex        =   15
      Top             =   1620
      Width           =   1335
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Modelo:"
      Height          =   195
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   555
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CONFIG. IMPRESSÃO DO CERTIFICADO"
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
      Width           =   8445
   End
End
Attribute VB_Name = "Form_CoordImpCert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim RsCoordImprCert As Recordset

Dim Modelo          As Integer
Dim tmp             As Integer
Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub
Private Sub CarregarCoord()
    Set RsCoordImprCert = BD.OpenRecordset("SELECT * FROM CoordImprCert WHERE Modelo = " & Modelo)
    If RsCoordImprCert.BOF And RsCoordImprCert.EOF Then
            MSFG_CoordCert.Rows = 1
            chkUsarCidRed.Value = 0
            chkUsarSigla.Value = 0
            MsgBox "Erro ao localizar coordenadas de impresssão.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsCoordImprCert.MoveFirst
            tmp = 0
            Do Until RsCoordImprCert.EOF
                tmp = tmp + 1
                MSFG_CoordCert.TextMatrix(tmp, 1) = RsCoordImprCert.Fields("Tp") & " mm"
                MSFG_CoordCert.TextMatrix(tmp, 2) = RsCoordImprCert.Fields("Me") & " mm"
                
                RsCoordImprCert.MoveNext
            Loop
            'For Tmp = 1 To 22
            '    Topo(Tmp) = IIf(IsNull(RsCoordImprCert.Fields("Tp(" & Tmp & ")")), 0, RsCoordImprCert.Fields("tp(" & Tmp & ")"))
            '    MargE(Tmp) = IIf(IsNull(RsCoordImprCert.Fields("ME(" & Tmp & ")")), 0, RsCoordImprCert.Fields("ME(" & Tmp & ")"))
            'Next
    End If
    'With MSFG_CoordCert
    '   For Tmp = 1 To 22
    '        .TextMatrix(Tmp, 1) = Topo(Tmp) & " mm"
    '        .TextMatrix(Tmp, 2) = MargE(Tmp) & " mm"
    '    Next
        
    Txt_Coord.Visible = False
    'End With
End Sub

Private Sub CarregarGrid()

    'If Trim(Cb_Modelo.Text) = "" Then Exit Sub
    Set RsCoordImprCert = BD.OpenRecordset("SELECT * FROM CoordImprCert WHERE Modelo = " & Modelo)
    If RsCoordImprCert.BOF And RsCoordImprCert.EOF Then
            MSFG_CoordCert.Rows = 1
            chkUsarCidRed.Value = 0
            chkUsarSigla.Value = 0
            MsgBox "Coordenadas do modelo " & Modelo & " não encontrado.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsCoordImprCert.MoveFirst
            
            With MSFG_CoordCert
                .Rows = 1
                Do Until RsCoordImprCert.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = RsCoordImprCert.Fields("campo") & " - " & RsCoordImprCert.Fields("Descr")
                    .TextMatrix(.Rows - 1, 1) = RsCoordImprCert.Fields("Tp") & " mm"
                    .TextMatrix(.Rows - 1, 2) = RsCoordImprCert.Fields("Me") & " mm"
                    RsCoordImprCert.MoveNext
                Loop
                .Col = 0
                .ColSel = 0
                .Row = 1
                .RowSel = .Rows - 1
                .FillStyle = flexFillRepeat
                .CellAlignment = 1
                .Row = 1
                .Col = 1
            End With
    End If
    Txt_Coord.Visible = False
    'Exit Sub
   ' With MSFG_CoordCert
   '     .Rows = 23
   '     .TextMatrix(1, 0) = "Nome da Instituição"
   '     .TextMatrix(2, 0) = "Endereço"
   '     .TextMatrix(3, 0) = "Criação da Inst. de Ensino"
   '     .TextMatrix(4, 0) = "Autorização do Curso"
   '     .TextMatrix(5, 0) = "Nome do C.E.S."
   '     .TextMatrix(6, 0) = "Nome do Aluno"
   '     .TextMatrix(7, 0) = "Nacionalidade"
   '     .TextMatrix(8, 0) = "Identidade"
   '     .TextMatrix(9, 0) = "Orgão Emissor"
   '     .TextMatrix(10, 0) = "Natural"
    '    .TextMatrix(11, 0) = "UF"
   '     .TextMatrix(12, 0) = "Data Nascimento DIA"
   '     .TextMatrix(13, 0) = "Data Nascimento MÊS"
   '     .TextMatrix(14, 0) = "Data Nascimento ANO"
   '     .TextMatrix(15, 0) = "Data Conclusão DIA"
   '     .TextMatrix(16, 0) = "Data Conclusão MÊS"
   '     .TextMatrix(17, 0) = "Data Conclusão ANO"
   '     .TextMatrix(18, 0) = "Ensino"
   '     .TextMatrix(19, 0) = "Municipio"
   '     .TextMatrix(20, 0) = "DIA Atual"
   '     .TextMatrix(21, 0) = "MÊS Atual"
   '     .TextMatrix(22, 0) = "ANO Atual"
   '     For Tmp = 0 To 1000
   '         Txt_Coord.AddItem (Tmp & " mm")
   '     Next
   '     Select Case modelo
   '         Case 1
   '             .Rows = 34
   '             .TextMatrix(23, 0) = "Verso: Nome do Aluno"
   '             .TextMatrix(24, 0) = "Verso: Disciplina"
    '            .TextMatrix(25, 0) = "Verso: Estab. de Ensino"
    '            .TextMatrix(26, 0) = "Verso: Cidade"
   '             .TextMatrix(27, 0) = "Verso: Estado"
   '             .TextMatrix(28, 0) = "Verso: Mês"
   '             .TextMatrix(29, 0) = "Verso: Ano"
   '             .TextMatrix(30, 0) = "Verso: Situação Final"
   '             .TextMatrix(31, 0) = "Verso: DIA do Cert."
   '             .TextMatrix(32, 0) = "Verso: Mês do Cert."
   '             .TextMatrix(33, 0) = "Verso: Ano do Cert."
    '            .Col = 0
   '             .ColSel = 0
   '             .Row = 23
   '             .RowSel = .Rows - 1
   '             .FillStyle = flexFillRepeat
   '             .CellBackColor = vbGreen
   '             '.CellAlignment = 1
   '             '.Row = 1
   '             '.Col = 1

    '        Case 2
    '            .Rows = 46
    '            .TextMatrix(23, 0) = "Verso: Disciplina"
    '            .TextMatrix(24, 0) = "Verso: Forma de Estudo"
    '            .TextMatrix(25, 0) = "Verso: Data"
    '            .TextMatrix(26, 0) = "Verso: Total de Horas"
     '           .TextMatrix(27, 0) = "Verso: Estab. de Ensino"
    '            .TextMatrix(28, 0) = "Verso: Cidade"
    '            .TextMatrix(29, 0) = "Verso: Estado"
    '            .TextMatrix(30, 0) = "Verso: Mês"
    '            .TextMatrix(31, 0) = "Verso: Ano"
     '           .TextMatrix(32, 0) = "Verso: Situação Final"
     '           .TextMatrix(33, 0) = "Verso: Curso Anterior"
    '            .TextMatrix(34, 0) = "Verso: Estabelecimento"
    '            .TextMatrix(35, 0) = "Verso: Localidade UF"
    '            .TextMatrix(36, 0) = "Verso: Outras Habilitações"
    '            .TextMatrix(37, 0) = "Verso: Outros"
    '            .TextMatrix(38, 0) = "Verso: Registro Num."
    '            .TextMatrix(39, 0) = "Verso: Folha"
    '            .TextMatrix(40, 0) = "Verso: Livro Num."
    '            .TextMatrix(41, 0) = "Verso: Publicação Dia"
    '            .TextMatrix(42, 0) = "Verso: Publicação Mês"
    '            .TextMatrix(43, 0) = "Verso: Publicação Ano"
    '            .TextMatrix(44, 0) = "Verso: Local"
    '            .TextMatrix(45, 0) = "Verso: Data"
    '
    '            .Col = 0
    '            .ColSel = 0
    '            .Row = 23
    '            .RowSel = .Rows - 1
    '            .FillStyle = flexFillRepeat
    '            .CellBackColor = vbGreen
    '        Case Else
    '            .Rows = 1
    '    End Select

        'CarregarCoord
    'End With
    
End Sub

Private Sub Bt_Gravar_Click()
    
    Set RsCoordImprCert = BD.OpenRecordset("SELECT * FROM CoordImprCert WHERE Modelo = " & Modelo & " ORDER BY Campo")
    If RsCoordImprCert.BOF And RsCoordImprCert.EOF Then
            MSFG_CoordCert.Rows = 1
            'MSFG_CoordCert.Rows = 2
            MsgBox "Erro ao localizar 'Modelo' de coordenadas de impressão.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            
            RsCoordImprCert.MoveFirst
            
            For tmp = 1 To MSFG_CoordCert.Rows - 1
                RsCoordImprCert.Edit '.AddNew
                'RsCoordImprCert.Fields("modelo") = modelo
                'RsCoordImprCert.Fields("Campo") = Tmp
                'RsCoordImprCert.Fields("Descr") = MSFG_CoordCert.TextMatrix(Tmp, 0)
                RsCoordImprCert.Fields("Tp") = Trim(left(MSFG_CoordCert.TextMatrix(tmp, 1), Len(MSFG_CoordCert.TextMatrix(tmp, 1)) - 2))
                RsCoordImprCert.Fields("Me") = Trim(left(MSFG_CoordCert.TextMatrix(tmp, 2), Len(MSFG_CoordCert.TextMatrix(tmp, 2)) - 2))
                RsCoordImprCert.Update
                RsCoordImprCert.MoveNext
            Next
            BD.Execute "UPDATE Ensino SET UsarCidReduzida = " & chkUsarCidRed.Value & ", UsarInstSigla = " & chkUsarSigla.Value & " WHERE ID = " & Modelo
            BD.Execute "UPDATE Config SET HB = '" & Trim(txtHB.Text) & "', NH = '" & Trim(txtNH.Text) & "', FormEstudo = '" & Trim(txtFormEstudo.Text) & "' WHERE Unidade = " & UnidadeEnsino
            
            
            
    End If
    Cb_Modelo.Clear
    chkUsarCidRed.Value = 0
    chkUsarSigla.Value = 0
    MSFG_CoordCert.Rows = 1
End Sub

Private Sub Cb_Modelo_Click()
    If Trim(Cb_Modelo.Text) = "" Then Exit Sub
    Modelo = left(Cb_Modelo.Text, 2)
    '...................................................................
   
    chkUsarCidRed.Value = IIf(pgUsarCidRed(Modelo) = True, 1, 0)
    chkUsarSigla.Value = IIf(pgUsarInstSigla(Modelo) = True, 1, 0)
    MostrarSitFinal
    
    '........................................................................
    CarregarGrid
End Sub

Private Sub Cb_Modelo_DropDown()
    Dim Rst As Recordset
    Cb_Modelo.Clear
    Set Rst = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY ID")
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                Cb_Modelo.AddItem left("00", 2 - Len(Rst.Fields("id"))) & Rst.Fields("id") & " - " & Rst.Fields("Descr")
                Rst.MoveNext
            Loop
    End If
    
    Rst.Close
    'Cb_Modelo.AddItem "01 - Ensino Fundamental"
    'Cb_Modelo.AddItem "02 - Ensino Médio"
End Sub

Private Sub Form_Load()
    'CarregarGrid
    
    MSFG_CoordCert.Rows = 1
    txtFormEstudo.Text = ""
End Sub




Private Sub MSFG_CoordCert_EnterCell()
        If MSFG_CoordCert.MouseCol = 0 Or MSFG_CoordCert.MouseRow = 0 Then
            Txt_Coord.Visible = False
            Exit Sub
        End If
        Txt_Coord.Visible = False
        Txt_Coord.top = MSFG_CoordCert.top + MSFG_CoordCert.CellTop
        Txt_Coord.left = MSFG_CoordCert.left + MSFG_CoordCert.CellLeft
        Txt_Coord.Width = MSFG_CoordCert.CellWidth
        Txt_Coord.Height = MSFG_CoordCert.CellHeight
        Txt_Coord.Text = Trim(Mid(MSFG_CoordCert.TextMatrix(MSFG_CoordCert.Row, MSFG_CoordCert.Col), 1, Len(MSFG_CoordCert.TextMatrix(MSFG_CoordCert.Row, MSFG_CoordCert.Col)) - 2))
        Txt_Coord.Visible = True
End Sub

Private Sub MSFG_CoordCert_LeaveCell()
    If MSFG_CoordCert.Col = 0 And MSFG_CoordCert.Row = 1 Or Txt_Coord.Visible = False Then
            'Txt_Coord.Visible = False
            Exit Sub
        End If
    MSFG_CoordCert.TextMatrix(MSFG_CoordCert.Row, MSFG_CoordCert.Col) = CInt(Txt_Coord.Text) & " mm"
End Sub

Private Sub MSFG_CoordCert_Scroll()
    MSFG_CoordCert_LeaveCell
    Txt_Coord.Visible = False
End Sub



Private Sub Txt_Coord_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii = 13 Then
        MSFG_CoordCert_LeaveCell
        Txt_Coord.Visible = False
        Exit Sub
    End If
    If KeyAscii >= 48 And KeyAscii <= 57 Then
        Else
            KeyAscii = 0
            Exit Sub
    End If
End Sub

Private Sub MostrarSitFinal()
    Dim Rst As Recordset
    
    Set Rst = BD.OpenRecordset("SELECT * FROM Config WHERE Unidade=" & UnidadeEnsino)
    If Rst.BOF And Rst.EOF Then
            txtHB.Text = ""
            txtNH.Text = ""
            txtFormEstudo.Text = ""
            MsgBox "Erro ao localizar dados de Situacao Final"
        Else
            Rst.MoveFirst
            txtHB.Text = IIf(IsNull(Rst.Fields("HB")), "", Rst.Fields("HB"))
            txtNH.Text = IIf(IsNull(Rst.Fields("NH")), "", Rst.Fields("NH"))
            txtFormEstudo.Text = IIf(IsNull(Rst.Fields("FormEstudo")), "", Rst.Fields("FormEstudo"))
    End If
    Rst.Close
End Sub


