VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_RelatOrientacao 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Orientações"
   ClientHeight    =   4305
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8775
   Icon            =   "Form_RelatOrientacao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4305
   ScaleWidth      =   8775
   Begin MSComDlg.CommonDialog cd 
      Left            =   4260
      Top             =   4020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkExpArquivo 
      Caption         =   "Exportar para arquivo"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Top             =   4020
      Width           =   3375
   End
   Begin VB.Frame Frame5 
      Height          =   1395
      Left            =   6180
      TabIndex        =   14
      Top             =   420
      Width           =   2535
      Begin VB.OptionButton Option1 
         Caption         =   "Analitico"
         Height          =   255
         Index           =   1
         Left            =   180
         TabIndex        =   16
         Top             =   780
         Width           =   2235
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Sintético"
         Height          =   255
         Index           =   0
         Left            =   180
         TabIndex        =   15
         Top             =   360
         Value           =   -1  'True
         Width           =   2235
      End
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   855
      Left            =   6240
      Picture         =   "Form_RelatOrientacao.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   3060
      Width           =   2415
   End
   Begin VB.Frame Frame4 
      Caption         =   "Ordenar por:"
      Height          =   795
      Left            =   60
      TabIndex        =   11
      Top             =   3120
      Width           =   6015
      Begin VB.ComboBox cbOrdem 
         Height          =   315
         Left            =   300
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   300
         Width           =   5535
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Professor:"
      Height          =   795
      Left            =   60
      TabIndex        =   9
      Top             =   1320
      Width           =   6015
      Begin VB.ComboBox cbProfessor 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   300
         Width           =   5415
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Disciplina:"
      Height          =   795
      Left            =   60
      TabIndex        =   7
      Top             =   2220
      Width           =   6015
      Begin VB.ComboBox cbDisciplina 
         Height          =   315
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   300
         Width           =   5475
      End
   End
   Begin VB.CommandButton btImprimir 
      Caption         =   "&Imprimir"
      Height          =   855
      Left            =   6240
      Picture         =   "Form_RelatOrientacao.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2160
      Width           =   2415
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo:"
      Height          =   795
      Left            =   60
      TabIndex        =   1
      Top             =   420
      Width           =   6060
      Begin MSComCtl2.DTPicker dtpFinal 
         Height          =   285
         Left            =   3150
         TabIndex        =   2
         Top             =   315
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   55705601
         CurrentDate     =   38538
      End
      Begin MSComCtl2.DTPicker dtpIni 
         Height          =   285
         Left            =   810
         TabIndex        =   3
         Top             =   315
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   55705601
         CurrentDate     =   38538
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   180
         TabIndex        =   5
         Top             =   360
         Width           =   555
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Final:"
         Height          =   195
         Left            =   2520
         TabIndex        =   4
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "RELATÓRIO DE ORIENTAÇÕES"
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
      Width           =   8805
   End
End
Attribute VB_Name = "Form_RelatOrientacao"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim DisciplinaID As Integer

Private Sub ExportarArquivo(sSQL As String, Tp As String)
    Dim RsTMP       As Recordset
    Dim caminho     As String
    Dim sDados      As String
    
    
    Set RsTMP = BD.OpenRecordset(sSQL)
    If RsTMP.BOF And RsTMP.EOF Then
            Exit Sub
        Else
            RsTMP.MoveFirst
    End If
    
    cd.DialogTitle = "Local e Nome do Arquivo?"
    cd.InitDir = App.path
    cd.FileName = "rel_Orientacao"
    cd.filter = "Excel |*.xls"
    'cd.Filter = "Todos | *.*"
    
    cd.ShowSave
    caminho = Trim(cd.FileName)
    If caminho = "" Then Exit Sub
    
    Do Until RsTMP.EOF
        Select Case Tp
            Case "S"
                sDados = RsTMP.Fields("Nome") & ";" & RsTMP.Fields("DtOrientacao") & ";" & _
                        RsTMP.Fields("Ensino.Descr") & ";" & RsTMP.Fields("ContarDeDisciplinaID")
            Case "A"
                sDados = RsTMP.Fields("Professores.Nome") & ";" & RsTMP.Fields("DtOrientacao") & ";" & _
                        RsTMP.Fields("Ensino.Descr") & ";" & RsTMP.Fields("Disciplina.Descr") & ";" & _
                        RsTMP.Fields("MatrID") & ";" & RsTMP.Fields("Matriculas.Nome")

        End Select
                 
        
        Call ExpArq(caminho, sDados)
        RsTMP.MoveNext
    Loop
    MsgBox "Arquivo exportado!", vbInformation, "CESNet - Aviso"
End Sub

Private Sub btImprimir_Click()
    Dim Criterio        As String
    Dim RsTMP           As Recordset
    Dim TotRegistros    As Integer
    Dim TotOrient       As Integer
        '**********************************************
        '************* SINTETICO
        '**********************************************

If Option1(0).Value = True Then
    Criterio = "SELECT MatriculaOrientacao.DtOrientacao, Professores.Nome, Ensino.Descr, Disciplina.Descr, Count(MatriculaOrientacao.DisciplinaID) AS ContarDeDisciplinaID" & IIf(cbProfessor.Text = "", "", ", MatriculaOrientacao.ProfOrientID") & IIf(cbDisciplina.Text = "", "", ", MatriculaOrientacao.DisciplinaID") & ", MatriculaOrientacao.EnsinoID " & _
                "FROM (Disciplina INNER JOIN (Ensino INNER JOIN MatriculaOrientacao ON Ensino.ID = MatriculaOrientacao.EnsinoID) ON Disciplina.ID = MatriculaOrientacao.DisciplinaID) INNER JOIN Professores ON MatriculaOrientacao.ProfOrientID = Professores.ProfID " & _
                "GROUP BY MatriculaOrientacao.DtOrientacao, Professores.Nome, Ensino.Descr, Disciplina.Descr, MatriculaOrientacao.EnsinoID" & IIf(cbProfessor.Text = "", "", ", MatriculaOrientacao.ProfOrientID") & IIf(cbDisciplina.Text = "", "", ", MatriculaOrientacao.DisciplinaID") & " " & _
                "HAVING (((MatriculaOrientacao.DtOrientacao) >= #" & Format(dtpIni.Value, "mm/dd/yyyy") & "# And (MatriculaOrientacao.DtOrientacao) <= #" & Format(dtpFinal.Value, "mm/dd/yyyy") & "#) " & _
                IIf(cbProfessor.Text = "", "", "AND ((MatriculaOrientacao.ProfOrientID)=" & PgIDProfessor(cbProfessor.Text) & ") ") & _
                IIf(cbDisciplina.Text = "", "", "AND ((MatriculaOrientacao.DisciplinaID)=" & PgIDDisciplina(cbDisciplina.Text) & ") ")
    
    'cbOrdem.AddItem "01 - DATA DE ORIENTAÇÃO"
    'cbOrdem.AddItem "02 - CURSO E DATA DE ORIENTAÇÃO"

    Select Case left(cbOrdem.Text, 2)
        Case "01" '01 - DATA DE ORIENTACAO
            Criterio = Criterio & ") ORDER BY MatriculaOrientacao.DtOrientacao"
        Case "02" '02 - CURSO E DATA DE ORIENTACAO
            Criterio = Criterio & ") ORDER BY MatriculaOrientacao.EnsinoID, MatriculaOrientacao.DtOrientacao"
        Case Else
            Criterio = Criterio & ") ORDER BY MatriculaOrientacao.DtOrientacao"
    End Select
           
    Set RsTMP = BD.OpenRecordset(Criterio)
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Nenhum Registro encontrado"
            Exit Sub
        Else
            RsTMP.MoveLast
            TotRegistros = RsTMP.RecordCount
            RsTMP.MoveFirst
            TotOrient = 0
            Do Until RsTMP.EOF
                TotOrient = TotOrient + RsTMP.Fields("ContarDeDisciplinaID")
                RsTMP.MoveNext
            Loop
            
    End If
    
    'EXPoRTAR ARQUIVO
            If chkExpArquivo.Value = 1 Then
                Call ExportarArquivo(Criterio, "S")
                Exit Sub
            End If
            
        
            rptOrientacaoS.Sections("Section2").Controls("lbTotalReg").Caption = TotRegistros
            rptOrientacaoS.Sections("Section2").Controls("lbTotOrient").Caption = TotOrient
            
            Call Relatorio(rptOrientacaoS, Criterio)
           'rptOrientacao.Sections("Corpo").Controls.Item("lbProfessor").Caption = cbProfessor.Text
            rptOrientacaoS.Show 1
            
            
        Else
        '**********************************************
        '************* ANALITICO
        '**********************************************
         Criterio = "SELECT Professores.Nome, Ensino.Descr, Disciplina.Descr, MatriculaOrientacao.DtOrientacao, Matriculas.MatrID, Matriculas.Nome, MatriculaOrientacao.ProfOrientID, MatriculaOrientacao.DisciplinaID " & _
                    "FROM (((MatriculaOrientacao INNER JOIN Disciplina ON MatriculaOrientacao.DisciplinaID = Disciplina.ID) INNER JOIN Ensino ON MatriculaOrientacao.EnsinoID = Ensino.ID) INNER JOIN Matriculas ON MatriculaOrientacao.MatrID = Matriculas.MatrID) INNER JOIN Professores ON MatriculaOrientacao.ProfOrientID = Professores.ProfID " & _
                    "WHERE (((MatriculaOrientacao.DtOrientacao)>=#" & Format(dtpIni.Value, "mm/dd/yyyy") & "# And (MatriculaOrientacao.DtOrientacao)<=#" & Format(dtpFinal.Value, "mm/dd/yyyy") & "#) " & _
                    IIf(cbProfessor.Text = "", "", " AND ((MatriculaOrientacao.ProfOrientID)=" & PgIDProfessor(cbProfessor.Text) & ") ") & _
                    IIf(cbDisciplina.Text = "", "", " AND ((MatriculaOrientacao.DisciplinaID)=" & PgIDDisciplina(cbDisciplina.Text) & ")")



    
    


    Select Case left(cbOrdem.Text, 2)
        Case "01" '01 - DATA DE ORIENTACAO
            Criterio = Criterio & ") ORDER BY MatriculaOrientacao.DtOrientacao"
        Case "02" '02 - CURSO E DATA DE ORIENTACAO
            Criterio = Criterio & ") ORDER BY MatriculaOrientacao.EnsinoID, MatriculaOrientacao.DtOrientacao"
        Case Else
            Criterio = Criterio & ") ORDER BY MatriculaOrientacao.DtOrientacao"
    End Select
           
    Set RsTMP = BD.OpenRecordset(Criterio)
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Nenhum Registro encontrado"
            Exit Sub
        Else
            RsTMP.MoveLast
            TotOrient = RsTMP.RecordCount
            RsTMP.Close
            
    End If
           
           
             'EXPoRTAR ARQUIVO
            If chkExpArquivo.Value = 1 Then
                Call ExportarArquivo(Criterio, "A")
                Exit Sub
            End If
            
            rptOrientacaoA.Sections("Section2").Controls("lbTotalReg").Caption = TotOrient 'TotRegistros
            rptOrientacaoA.Sections("Section2").Controls("lbTotOrient").Caption = TotOrient
            
            Call Relatorio(rptOrientacaoA, Criterio)
            rptOrientacaoA.Show 1

    End If
End Sub

Private Sub cbDisciplina_Click()
    If Trim(cbDisciplina.Text) = "" Then Exit Sub
    DisciplinaID = PgIDDisciplina(Trim(cbDisciplina.Text))
End Sub

Private Sub cbDisciplina_DropDown()
    Dim RsTMP As Recordset
    cbDisciplina.Clear
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Disciplina ORDER BY Descr")
    RsTMP.MoveFirst
    Do Until RsTMP.EOF
        cbDisciplina.AddItem RsTMP.Fields("Descr")
        RsTMP.MoveNext
    Loop
    RsTMP.Close
End Sub

Private Sub cbOrdem_DropDown()
    cbOrdem.Clear
    cbOrdem.AddItem "01 - DATA DE ORIENTAÇÃO"
    cbOrdem.AddItem "02 - CURSO E DATA DE ORIENTAÇÃO"
    
End Sub


Private Sub cbProfessor_DropDown()
    Dim RsTMP As Recordset
    cbProfessor.Clear
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Professores ORDER BY Nome")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
            Exit Sub
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                cbProfessor.AddItem RsTMP.Fields("Nome")
                RsTMP.MoveNext
            Loop
            RsTMP.Close
    End If
End Sub

Private Sub Command1_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    dtpIni.Value = Date - 30
    dtpFinal.Value = Date
End Sub

