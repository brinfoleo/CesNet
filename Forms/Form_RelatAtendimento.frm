VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_RelatAtendimentos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Relatorio de Atendimentos "
   ClientHeight    =   1860
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5910
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   1860
   ScaleWidth      =   5910
   Begin VB.Frame Frame1 
      Caption         =   "Periodo"
      Height          =   1230
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   3360
      Begin MSComCtl2.DTPicker DTP_Final 
         Height          =   285
         Left            =   810
         TabIndex        =   2
         Top             =   675
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56492033
         CurrentDate     =   38538
      End
      Begin MSComCtl2.DTPicker DTP_Ini 
         Height          =   285
         Left            =   810
         TabIndex        =   3
         Top             =   315
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   503
         _Version        =   393216
         Format          =   56492033
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
         Left            =   180
         TabIndex        =   4
         Top             =   720
         Width           =   555
      End
   End
   Begin VB.CommandButton Bt_Imprimir 
      Caption         =   "Imprimir"
      Height          =   795
      Left            =   3720
      Picture         =   "Form_RelatAtendimento.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   720
      Width           =   2025
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "RELATÓRIO DE ATENDIMENTOS DIVERSOS"
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
      TabIndex        =   6
      Top             =   0
      Width           =   6090
   End
End
Attribute VB_Name = "Form_RelatAtendimentos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit



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

Private Sub Bt_Imprimir_Click()
    If ChkAcesso(Me.Name, "I") = False Then Exit Sub
    
    Dim Criterio    As String
    Dim RsTMP       As Recordset
    Dim TotReg      As String
    
    'Criterio = "SELECT * FROM RegAtendimento WHERE Dt BETWEEN #" & Format(DTP_Ini.Value, "mm/dd/yyyy") & "# AND #" & Format(DTP_Final.Value, "mm/dd/yyyy") & "#"
    'Criterio = "SELECT RegAtendimento.*, Usuario.Responsavel " & _
                "FROM Usuario INNER JOIN RegAtendimento ON Usuario.UsuarioID = RegAtendimento.UsuID"

    Criterio = "SELECT RegAtendimento.id, Usuario.Responsavel, RegAtendimento.Dt, RegAtendimento.UsuID, RegAtendimento.Hr, RegAtendimento.Motivo" & _
                " FROM Usuario INNER JOIN RegAtendimento ON Usuario.UsuarioID = RegAtendimento.UsuID " & _
                " WHERE (((RegAtendimento.Dt)>=#" & Format(DTP_Ini.Value, "mm/dd/yyyy") & "# And (RegAtendimento.Dt)<=#" & Format(DTP_Final.Value, "mm/dd/yyyy") & "#))"

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
    'If chkExpArquivo.Value = 1 Then
    '    Call ExportarArquivo(Criterio, "S")
    '    Exit Sub
    'End If
    '====================================
    'Criterio = "SELECT * FROM MatriculaEnsino WHERE DtInicio >= #" & Format(DTP_Ini.Value, "mm/dd/yyyy") & "# AND DtInicio <= #" & Format(DTP_Final.Value, "mm/dd/yyyy") & "# ORDER BY MatrID"
    'lbTotalReg
    rptListAtendDiv.Sections("Rodape").Controls.Item("lbTotalReg").Caption = "Total de Registros: " & TotReg
    
    Call Relatorio(rptListAtendDiv, Criterio)
    rptListAtendDiv.Show 1
    
End Sub


