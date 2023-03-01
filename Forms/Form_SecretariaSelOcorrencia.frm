VERSION 5.00
Begin VB.Form Form_SecretariaSelOcorrencia 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Selecione a Ocorrencia"
   ClientHeight    =   1755
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7830
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1755
   ScaleWidth      =   7830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Bt_Cancelar 
      Cancel          =   -1  'True
      Caption         =   "&Cancelar"
      Height          =   795
      Left            =   5880
      Picture         =   "Form_SecretariaSelOcorrencia.frx":0000
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   900
      Width           =   1755
   End
   Begin VB.CommandButton Bt_Aplicar 
      Caption         =   "&Aplicar"
      Height          =   795
      Left            =   3960
      Picture         =   "Form_SecretariaSelOcorrencia.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   900
      Width           =   1755
   End
   Begin VB.ComboBox Cb_Ocorrencia 
      Height          =   315
      Left            =   1260
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Ocorrencia:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "Form_SecretariaSelOcorrencia"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim OcorrenciaID    As Integer

Private Sub Bt_Aplicar_Click()
    If Trim(Cb_Ocorrencia.Text) = "" Then
            OcorrenciaID = 0
        Else
            OcorrenciaID = Left(Trim(Cb_Ocorrencia.Text), 3)
    End If
    Unload Me
End Sub

Private Sub Bt_Cancelar_Click()
    OcorrenciaID = 0
    Unload Me
End Sub

Private Sub Cb_Ocorrencia_DropDown()
    'If Trim(MatrID) = "" Then Exit Sub
    Dim RsOcorrencia As Recordset
    Cb_Ocorrencia.Clear
    Set RsOcorrencia = BD.OpenRecordset("SELECT * FROM OcorrenciaConclusao ORDER BY Descr ASC")
    If RsOcorrencia.BOF And RsOcorrencia.EOF Then
        Else
            RsOcorrencia.MoveFirst
            Do Until RsOcorrencia.EOF
                Cb_Ocorrencia.AddItem Left("000", 3 - Len(Trim(RsOcorrencia.Fields("OcorrenciaID")))) & RsOcorrencia.Fields("OcorrenciaID") & " - " & RsOcorrencia.Fields("Descr")
                RsOcorrencia.MoveNext
            Loop
            
    End If
    RsOcorrencia.Close
End Sub


Public Function SelOcorrencia() As Integer
    Me.Show 1
    
    

    SelOcorrencia = OcorrenciaID
    'Unload Me
End Function

