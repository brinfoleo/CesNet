VERSION 5.00
Begin VB.Form Form_ImpressaoCertificado 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Impressão de Certificado"
   ClientHeight    =   2040
   ClientLeft      =   270
   ClientTop       =   495
   ClientWidth     =   6420
   Icon            =   "Form_ImpressaoCertificado.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2040
   ScaleWidth      =   6420
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   750
      Left            =   3960
      Picture         =   "Form_ImpressaoCertificado.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   975
      Width           =   2310
   End
   Begin VB.CommandButton Bt_Imprimir 
      Caption         =   "&Imprimir"
      Height          =   750
      Left            =   3900
      Picture         =   "Form_ImpressaoCertificado.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   120
      Width           =   2310
   End
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   3705
      Begin VB.OptionButton Option1 
         Caption         =   "Imprimir a VERSO do Certificado."
         Height          =   330
         Index           =   1
         Left            =   585
         TabIndex        =   2
         Top             =   900
         Width           =   2850
      End
      Begin VB.OptionButton Option1 
         Caption         =   "Imprimir a FRENTE do Certificado."
         Height          =   330
         Index           =   0
         Left            =   585
         TabIndex        =   1
         Top             =   450
         Width           =   2760
      End
   End
End
Attribute VB_Name = "Form_ImpressaoCertificado"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LadoCert        As Integer
Dim Modelo          As Integer
Dim DtCertificado   As String
Dim Matricula       As String
Public Function ImprimirCertificado(nModelo As Integer, MatrID As String, DtCert As String)
    Modelo = nModelo
    DtCertificado = DtCert
    Matricula = MatrID
    Me.Show 1
End Function

Private Sub Bt_Cancelar_Click()
    Unload Me
End Sub

Private Sub Bt_Imprimir_Click()
Call ImprCertificado(LadoCert, Modelo, Matricula, DtCertificado)
End Sub

Private Sub Option1_Click(Index As Integer)
    LadoCert = Index
End Sub
