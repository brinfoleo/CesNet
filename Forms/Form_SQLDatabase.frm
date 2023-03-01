VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "msdatgrd.ocx"
Begin VB.Form Form_SQLDatabase 
   Caption         =   "SQL Database"
   ClientHeight    =   5925
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9750
   LinkTopic       =   "Form1"
   ScaleHeight     =   5925
   ScaleWidth      =   9750
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Eliminar Cursos em Duplicidade"
      Height          =   495
      Left            =   8040
      TabIndex        =   7
      Top             =   240
      Width           =   1395
   End
   Begin VB.CommandButton Command1 
      Caption         =   ">>"
      Height          =   255
      Left            =   7200
      TabIndex        =   6
      Top             =   480
      Width           =   435
   End
   Begin VB.TextBox txtSQL 
      Height          =   285
      Left            =   720
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   480
      Width           =   6315
   End
   Begin VB.Frame frmTab 
      Caption         =   "Tabela:"
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
      Left            =   120
      TabIndex        =   2
      Top             =   1140
      Width           =   9315
      Begin MSDataGridLib.DataGrid DBGrid 
         Height          =   4395
         Left            =   120
         TabIndex        =   3
         Top             =   300
         Width           =   8475
         _ExtentX        =   14949
         _ExtentY        =   7752
         _Version        =   393216
         AllowUpdate     =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         AllowAddNew     =   -1  'True
         AllowDelete     =   -1  'True
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
   Begin VB.ComboBox cboTabelas 
      Height          =   315
      Left            =   780
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   60
      Width           =   6555
   End
   Begin VB.Label Label2 
      Caption         =   "SQL:"
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   480
      Width           =   435
   End
   Begin VB.Label Label1 
      Caption         =   "Tabelas"
      Height          =   315
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   795
   End
End
Attribute VB_Name = "Form_SQLDatabase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Sql             As String



Private Sub cboTabelas_Click()
    Sql = ""
    openTabela
End Sub

Private Sub Command1_Click()
    
    Sql = txtSQL.Text
        
    openTabela
End Sub




Private Sub Command2_Click()
    On Error GoTo TrErro
    Me.Enabled = False
    EliminarDuplicidadeEnsino
    Me.Enabled = True
    Exit Sub
TrErro:
    Me.Enabled = True
End Sub

Private Sub Form_Activate()
    'If InputBox("Digite a senha:", "Manutenção") <> "131211" Then
    '    Unload Me
    'End If
End Sub

Private Sub Form_Load()
    
    cboTabelas.Clear
    txtSQL = ""
    Dim rstSchema As ADODB.Recordset
    Dim strCnn As String
    AbrirBD_ADO
    Set rstSchema = BD_ADO.OpenSchema(adSchemaTables)
    Do Until rstSchema.EOF
        cboTabelas.AddItem (rstSchema!TABLE_NAME)
        rstSchema.MoveNext
    Loop
    rstSchema.Close
End Sub

Private Sub openTabela()
    On Error GoTo TrtERROGrid
    Dim nome_tabela     As String
    
    Dim Tabela          As String
    Dim Rst             As ADODB.Recordset

    Tabela = cboTabelas.Text
    frmTab.Caption = "Tabela: " & Tabela
    
    
    If Trim(Sql) = "" Then
        Sql = "SELECT * FROM " & Tabela
    End If

    
    Set Rst = New ADODB.Recordset
    Rst.CursorLocation = adUseClient
    Rst.Open Sql, BD_ADO ', adOpenForwardOnly, adLockPessimistic
    
    
    'Set Rst = BD_ADO.Open(sql)
    
    
    
    If Rst Is Nothing Then
        'DataGrid1.Enabled = False
        'Text1.Enabled = False
        Me.Caption = "Busca - [ 00000 Registros]"
        Exit Sub
    End If
    
    If Rst.BOF And Rst.EOF Then
            'DataGrid1.Enabled = False
            'Text1.Enabled = False
            Me.Caption = "Busca - [ 00000 Registros]"
        Else
            DBGrid.Enabled = True
            'Rst.MoveLast
            Me.Caption = "Busca - [ " & left(String(5, "0"), 5 - Len(Trim(Rst.RecordCount))) & Trim(Rst.RecordCount) & " Registros]"
            Rst.MoveFirst
    End If
    Set DBGrid.DataSource = Rst.DataSource
    With DBGrid
        .AllowUpdate = False
        '.EditActive = False
        '.Columns(0).Caption = "Razão Social"
        ''.Columns(0).DataField = Rst.Fields("xNome")
        '.Columns(0).Width = TextWidth("Razão Social") + 1000
        '.Columns(0).Alignment = dbgLeft

        '.Columns(1).Caption = "Razão Social"
        ''.Columns(1).DataField = Rst.Fields("xNome")
        '.Columns(1).Width = TextWidth("Razão Social") + 1000
        '.Columns(1).Alignment = dbgLeft
        
    End With
    Exit Sub
TrtERROGrid:
    MsgBox "Descrição: " & Err.Description, vbCritical, "Erro n. " & Err.Number
    Sql = ""
    'resultadoBusca = ""
    'Unload Me
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    With frmTab
        .top = 950
        .left = 150
        .Width = Me.ScaleWidth - 300
        .Height = Me.ScaleHeight - (.top + 150)
    End With
    With DBGrid
        .Height = frmTab.Height - 400
        .Width = frmTab.Width - 250
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    BD_ADO.Close
End Sub


Private Sub txtSQL_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        Command1_Click
    End If
End Sub
