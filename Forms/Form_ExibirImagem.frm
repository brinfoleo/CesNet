VERSION 5.00
Begin VB.Form Form_ExibirImagem 
   Caption         =   "CESNet - Imagem"
   ClientHeight    =   3540
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4185
   Icon            =   "Form_ExibirImagem.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3540
   ScaleWidth      =   4185
   Begin VB.Label lbMatricula 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
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
      Width           =   6705
   End
   Begin VB.Image foto 
      Height          =   2415
      Left            =   360
      Stretch         =   -1  'True
      Top             =   840
      Width           =   2235
   End
End
Attribute VB_Name = "Form_ExibirImagem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim nFoto As String


Public Sub ExibirFoto(MatrID As String)
    If MatrID = "" Then
        MsgBox "Favor selecionar uma Matrícula!", vbInformation, "CESNet - Aviso"
        Unload Me
        Exit Sub
    End If
    nFoto = Format(MatrID, "000000000")
    lbMatricula.Caption = MatrID
    ExibeImagem
    
    'Me.Show
   
    
End Sub
Private Sub ExibeImagem()
 On Error GoTo Err_LPI
    Dim sImage As String
    
    DoEvents
    sImage = PathBD & "\Database\IMG\" & nFoto & "001.jpg"
    'foto.Picture = LoadPicture(ArqFoto)
    
    
    'Load the file pointer into the Image control
    If Len(Dir(sImage)) Then
            foto.Picture = LoadPicture(sImage)
        Else
            foto.Picture = LoadPicture()
    End If
    
    Exit Sub
Err_LPI:
    MsgBox Err.Source & vbCrLf & vbCrLf & Err.Description, vbInformation, Err.Number
    Err.Clear
End Sub

Private Sub Form_Load()
    Me.top = 10
    Me.left = 10
    Me.Height = 4475 'foto.Height
    Me.Width = 4330 'foto.Height
End Sub

Private Sub Form_Resize()
    On Error GoTo TrtErroSize
    ' Me.Height = IIf(Me.Height - 700 - lbMatricula.Height < 2100, 2100, Me.Height - 700 - lbMatricula.Height)
    lbMatricula.Width = Me.Width
    foto.top = 100 + lbMatricula.Height
    foto.left = 100
    foto.Width = Me.Width - 350
    foto.Height = IIf(Me.Height - 700 - lbMatricula.Height < 950, 950, Me.Height - 700 - lbMatricula.Height)
   
    Me.Caption = "CESNet - imagem (" & foto.Height & " x " & foto.Width & ")"
    Exit Sub
TrtErroSize:
    Resume Next
End Sub
