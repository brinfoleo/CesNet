VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_RelatAlunosCadastrados 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Listagem de Alunos Cadastrados"
   ClientHeight    =   2010
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6030
   Icon            =   "Form_RelatAlunosCadastrados.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2010
   ScaleWidth      =   6030
   Begin MSComDlg.CommonDialog cd 
      Left            =   3840
      Top             =   1500
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox chkExpArquivo 
      Caption         =   "Exportar para arquivo"
      Height          =   195
      Left            =   180
      TabIndex        =   7
      Top             =   1740
      Width           =   3375
   End
   Begin VB.CommandButton Bt_Imprimir 
      Caption         =   "Imprimir"
      Height          =   795
      Left            =   3660
      Picture         =   "Form_RelatAlunosCadastrados.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   660
      Width           =   2205
   End
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   1230
      Left            =   90
      TabIndex        =   1
      Top             =   450
      Width           =   3360
      Begin MSComCtl2.DTPicker DTP_Final 
         Height          =   285
         Left            =   810
         TabIndex        =   5
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   38538
      End
      Begin MSComCtl2.DTPicker DTP_Ini 
         Height          =   285
         Left            =   810
         TabIndex        =   4
         Top             =   315
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56033281
         CurrentDate     =   38538
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Final:"
         Height          =   195
         Left            =   180
         TabIndex        =   3
         Top             =   720
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Inicio:"
         Height          =   195
         Left            =   180
         TabIndex        =   2
         Top             =   360
         Width           =   555
      End
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "RELATÓRIO DE ALUNOS CADASTRADOS"
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
      Width           =   6090
   End
End
Attribute VB_Name = "Form_RelatAlunosCadastrados"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Bt_Imprimir_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    
    Dim Criterio    As String
    Dim RsTMP       As Recordset
    Dim TotReg      As String
    
    Criterio = "SELECT * " & _
                "FROM Matriculas  " & _
                "Where DtMat >= #" & Format(DTP_Ini.Value, "mm/dd/yyyy") & "# And dtMat <= #" & Format(DTP_Final.Value, "mm/dd/yyyy") & "# " & _
                "ORDER BY MatrID"
    Set RsTMP = BD.OpenRecordset(Criterio)
    If RsTMP.BOF And RsTMP.EOF Then
            MsgBox "Nenhum registro encontrado", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsTMP.MoveLast
            TotReg = RsTMP.RecordCount
    End If
    RsTMP.Close
    'Exportar arquivo======================
    If chkExpArquivo.Value = 1 Then
        Call ExportarArquivo(Criterio, "S")
        Exit Sub
    End If
    '====================================
    'Criterio = "SELECT * FROM MatriculaEnsino WHERE DtInicio >= #" & Format(DTP_Ini.Value, "mm/dd/yyyy") & "# AND DtInicio <= #" & Format(DTP_Final.Value, "mm/dd/yyyy") & "# ORDER BY MatrID"
    'lbTotalReg
    rptListAlunosCadastrados.Sections("Rodape").Controls.Item("lbTotalReg").Caption = "Total de Registros: " & TotReg
    
    Call Relatorio(rptListAlunosCadastrados, Criterio)
    rptListAlunosCadastrados.Show 1
    
End Sub
Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    DTP_Final.Value = Date
    DTP_Ini.Value = DTP_Final.Value - 30
End Sub

Private Sub ExportarArquivo(sSQL As String, Tp As String)
    Dim RsTMP       As Recordset
    Dim caminho     As String
    Dim sDados      As String
    
    Dim Defi As String
    
    
    Set RsTMP = BD.OpenRecordset(sSQL)
    If RsTMP.BOF And RsTMP.EOF Then
            Exit Sub
        Else
            RsTMP.MoveFirst
    End If
    
    cd.DialogTitle = "Local e Nome do Arquivo?"
    cd.InitDir = App.path
    cd.FileName = "rel_ListMatriculas"
    cd.filter = "Excel |*.xls"
    'cd.Filter = "Todos | *.*"
    
    cd.ShowSave
    caminho = Trim(cd.FileName)
    If caminho = "" Then Exit Sub
    
    Do Until RsTMP.EOF
        If IsNull(RsTMP.Fields("DefID")) = False Then
                    Defi = PgNomeDef(RsTMP.Fields("DefID"))
                Else
                    Defi = " "
        End If
                 
        sDados = IIf(IsNull(RsTMP.Fields("Descr")), " ", RsTMP.Fields("Descr")) & vbTab & _
                 IIf(IsNull(RsTMP.Fields("DtMat")), " ", RsTMP.Fields("DtMat")) & vbTab & _
                 IIf(IsNull(RsTMP.Fields("MatrID")), " ", RsTMP.Fields("MatrID")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Nome")), " ", RsTMP.Fields("Nome")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("End")), " ", RsTMP.Fields("End")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Numero")), " ", RsTMP.Fields("Numero")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Compl")), " ", RsTMP.Fields("Compl")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Bai")), " ", RsTMP.Fields("Bai")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Mun")), " ", RsTMP.Fields("Mun")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("UF")), " ", RsTMP.Fields("UF")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("CEP")), " ", RsTMP.Fields("CEP")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Sexo")), " ", RsTMP.Fields("Sexo")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Nasc")), " ", RsTMP.Fields("Nasc")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Mail")), " ", RsTMP.Fields("Mail")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Cel")), " ", RsTMP.Fields("Cel")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Tel1")), " ", RsTMP.Fields("Tel1")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Tel2")), " ", RsTMP.Fields("Tel2")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("CPF")), " ", RsTMP.Fields("CPF")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("RG")), " ", RsTMP.Fields("RG")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("OE")), " ", RsTMP.Fields("OE")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("CertNasc")), " ", RsTMP.Fields("CertNasc")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Natural")), " ", RsTMP.Fields("Natural")) & vbTab
        sDados = sDados & IIf(IsNull(RsTMP.Fields("NaturalUF")), " ", RsTMP.Fields("NaturalUF")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("EstCivil")), " ", RsTMP.Fields("EstCivil")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Nacion")), " ", RsTMP.Fields("Nacion")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Mae")), " ", RsTMP.Fields("Mae")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Pai")), " ", RsTMP.Fields("Pai")) & vbTab & _
                        Defi & vbTab & _
                        IIf(IsNull(RsTMP.Fields("Obs")), " ", RsTMP.Fields("Obs")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("ValCard")), " ", RsTMP.Fields("ValCard")) & vbTab & _
                        IIf(IsNull(RsTMP.Fields("NumAnt")), " ", RsTMP.Fields("NumAnt"))
                        
        Call ExpArq(caminho, sDados)
        RsTMP.MoveNext
    Loop
    MsgBox "Arquivo exportado!", vbInformation, "CESNet - Aviso"
End Sub




