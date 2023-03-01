VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form formBusca 
   Caption         =   "Busca"
   ClientHeight    =   6390
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   6390
   ScaleWidth      =   8805
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      Height          =   1515
      Left            =   60
      TabIndex        =   1
      Top             =   4800
      Width           =   8655
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Text            =   "Text1"
         Top             =   540
         Width           =   6855
      End
      Begin VB.CommandButton btAplicar 
         Caption         =   "&Aplicar"
         Height          =   435
         Left            =   5160
         TabIndex        =   3
         Top             =   960
         Width           =   1635
      End
      Begin VB.CommandButton btCancelar 
         Cancel          =   -1  'True
         Caption         =   "&Cancelar"
         Height          =   435
         Left            =   6900
         TabIndex        =   2
         Top             =   960
         Width           =   1635
      End
      Begin VB.Label Label2 
         Caption         =   "Campo de busca:"
         Height          =   195
         Left            =   60
         TabIndex        =   7
         Top             =   240
         Width           =   1275
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Texto de Busca:"
         Height          =   195
         Left            =   60
         TabIndex        =   6
         Top             =   600
         Width           =   1215
      End
      Begin VB.Label lbCampoBusca 
         Caption         =   "Label2"
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
         Left            =   1500
         TabIndex        =   5
         Top             =   240
         Width           =   3975
      End
   End
   Begin MSDataGridLib.DataGrid grdBusca 
      Height          =   4695
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   8281
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1046
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FormBusca"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'Dim RsADO               As ADODB.Recordset
Dim Rst                 As ADODB.Recordset 'ADODB.Recordset
Dim Tabela              As String
Dim campoBusca          As String
Dim CamposBusca         As String
Dim resultadoBusca      As String
Dim sizeColuna(100)     As Integer
Dim tpConex             As String 'Seleciona o tipo de Conexao: 'local' ao BD

Public Function IniciarBusca(nomeTabela As String, _
                             Optional sColunas As String, _
                             Optional DefaultCampoBusca As String, _
                             Optional DefaultTextoBusca As String, _
                             Optional tipoConexao = "local") As String
    On Error Resume Next
    tpConex = tipoConexao
    If nomeTabela = "" Then Exit Function
    If Trim(sColunas) = "" Then
            Select Case nomeTabela
                'Case "Matriculas"
                    'CamposBusca = "MatrId," & "Nome, Referencia, CodigoBarras, NCM,IPIAliquota,ICMSCST,saldo"
                Case Else
                    CamposBusca = "*"
                    'MsgBox "sem paramentros"
            End Select
        Else
            CamposBusca = IIf(Trim(sColunas) = "", "*", "Id," & sColunas)
    End If
    'Busca Default
    If Trim(DefaultCampoBusca) <> "" Then
        campoBusca = DefaultCampoBusca
        lbCampoBusca.Caption = DefaultCampoBusca
        Text1.Text = DefaultTextoBusca
    End If
    Tabela = nomeTabela
    resultadoBusca = 0
    FormBusca.Show 1
    IniciarBusca = resultadoBusca

End Function

Private Sub btAplicar_Click()
    Unload Me
End Sub

Private Sub btCancelar_Click()
    resultadoBusca = 0
    Unload Me
End Sub



Private Sub Form_Load()
    If tpConex = "local" Then
        AbrirBD_ADO
    End If
    Text1.Text = ""
    PreencherGrid
    campoBusca = grdBusca.Columns(1).Caption
    lbCampoBusca.Caption = campoBusca
End Sub

Private Sub PreencherGrid(Optional strSQL As String)
    On Error Resume Next
    Dim sSQL As String
    Dim i As Integer
    
    Set Rst = New ADODB.Recordset
    sSQL = "SELECT " & CamposBusca & " FROM " & Tabela & IIf(Trim(strSQL) = "", "", " " & strSQL)
    If tpConex = "local" Then
            Rst.Open sSQL, BD_ADO
            BD_ADO.CursorLocation = adUseClient
        Else
            Rst.Open sSQL, conexao
            conexao.CursorLocation = adUseClient
    End If
    
    
    If Rst Is Nothing Then
        grdBusca.Enabled = False
        Text1.Enabled = False
        Me.Caption = "Busca - [ 00000 Registros]"
        Exit Sub
    End If
    
    If Rst.BOF And Rst.EOF Then
            grdBusca.Enabled = False
            'Text1.Enabled = False
            Me.Caption = "Busca - [ 00000 Registros]"
            'Exit Sub
        Else
            grdBusca.Enabled = True
            Rst.MoveLast
            Me.Caption = "Busca - [ " & left(String(5, "0"), 5 - Len(Trim(Rst.RecordCount))) & Trim(Rst.RecordCount) & " Registros]"
            Rst.MoveFirst
             
    End If
   
    Set grdBusca.DataSource = Rst.DataSource
    

    With grdBusca
        .AllowUpdate = False
        For i = 0 To .Columns.Count - 1
            If .Columns(i).Caption = "Id_Empresa" Then .Columns(i).Visible = False
            If .Columns(i).Caption = "DtHr" Then .Columns(i).Visible = False
        Next
    End With
    For i = 1 To grdBusca.Columns.Count - 1
        grdBusca.Columns(i).Width = IIf(sizeColuna(i) = 0, grdBusca.Columns(i).Width, sizeColuna(i))
    Next
    Exit Sub

End Sub

Private Sub Form_Resize()
    On Error Resume Next
    grdBusca.Width = Me.ScaleWidth - 150
    grdBusca.Height = Me.ScaleHeight - (150 + Frame1.Height)
    
    Frame1.top = grdBusca.Height + 100
    
    Me.Width = IIf(Me.Width < 8925, 8925, Me.Width)
    Me.Height = IIf(Me.Height < 6900, 6900, Me.Height)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If tpConex = "local" Then
        BD_ADO.Close
    End If
End Sub

Private Sub grdBusca_Click()
    resultadoBusca = grdBusca.Columns(0).Text
    Text1.SetFocus
End Sub

Private Sub grdBusca_ColResize(ByVal ColIndex As Integer, Cancel As Integer)
    sizeColuna(ColIndex) = grdBusca.Columns(ColIndex).Width
End Sub

Private Sub grdBusca_DblClick()
    btAplicar_Click
End Sub

Private Sub grdBusca_HeadClick(ByVal ColIndex As Integer)
    campoBusca = grdBusca.Columns(ColIndex).Caption
    lbCampoBusca.Caption = campoBusca
    '******* Pega o Tipo de Campo *****************
    'Dim tipoCampo           As Integer
    'tipoCampo = Rst.Fields(campoBusca).Type
    '**********************************************
End Sub

Private Sub Text1_Change()
    Dim sSQL    As String
    Dim sBusca  As String
    Dim sBTMP   As String
    If Trim(Text1.Text) = "" Then
            PreencherGrid
        Else
        '"WHERE Descricao LIKE '%cobre%' AND Descricao Like '%ba%' ORDER BY Descricao"
            sBTMP = ""
            sBusca = Replace(Trim(Text1.Text), " ", "|") & "|"
            
            Do Until InStr(sBusca, "|") = 0
                sBTMP = IIf(Trim(sBTMP) = "", "", sBTMP & " AND ") & campoBusca & " LIKE '%" & Trim(Mid(sBusca, 1, InStr(sBusca, "|") - 1)) & "%'"
                sBusca = Mid(sBusca, InStr(sBusca, "|") + 1, Len(sBusca))
            Loop
            'sSQL = "WHERE " & campoBusca & " LIKE '%" & Replace(Trim(Text1.Text), " ", "%") & "%' ORDER BY " & campoBusca
            sSQL = "WHERE " & sBTMP & " ORDER BY " & campoBusca
            PreencherGrid (sSQL)
    End If
End Sub
