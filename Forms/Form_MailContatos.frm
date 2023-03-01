VERSION 5.00
Begin VB.Form Form_MailContatos 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Mail Contatos"
   ClientHeight    =   5850
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5850
   ScaleWidth      =   4455
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lst 
      Height          =   4560
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   1
      Top             =   420
      Width           =   4215
   End
   Begin VB.CommandButton btAplicar 
      Caption         =   "&Aplicar"
      Height          =   435
      Left            =   2940
      TabIndex        =   0
      Top             =   5340
      Width           =   1395
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CONTATOS"
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
      TabIndex        =   2
      Top             =   0
      Width           =   4485
   End
End
Attribute VB_Name = "Form_MailContatos"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Selecao As String
Private Sub btAplicar_Click()
    Dim i As Integer
    Selecao = ""
    For i = 0 To lst.ListCount - 1
        If lst.Selected(i) = True Then
            'If Trim(Selecao) = "" Then
                    'Selecao = Trim(left(lst.List(i), 3))
                'Else
                    'Selecao = Selecao & "(" & Trim(left(lst.List(i), 3)) & ")"
                    Selecao = Selecao & "(" & lst.List(i) & ")"
            'End If
        End If
    Next
    Unload Me
End Sub

Private Sub Form_Load()
    Dim RsTMP As Recordset
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Usuario ORDER BY Responsavel")
    If RsTMP.BOF And RsTMP.EOF Then
            lst.Clear
        Else
            RsTMP.MoveFirst
            Do Until RsTMP.EOF
                lst.AddItem RsTMP.Fields("Responsavel")  'left("000", 3 - Len(Trim(RsTMP.Fields("UsuarioID")))) & RsTMP.Fields("UsuarioID") & " - " & RsTMP.Fields("Responsavel")
                RsTMP.MoveNext
            Loop
    End If
        
End Sub
Public Function PgUsuContatos() As String
    Me.Show 1
    PgUsuContatos = Selecao
    
    Unload Me
End Function
