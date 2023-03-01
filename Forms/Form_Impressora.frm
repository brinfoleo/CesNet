VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_Impressora 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Configurar Impressão"
   ClientHeight    =   3120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7635
   Icon            =   "Form_Impressora.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3120
   ScaleWidth      =   7635
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList IL_ConfImp 
      Left            =   4680
      Top             =   2340
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.CommandButton Bt_Visualizar 
      Caption         =   "Visualizar"
      Height          =   825
      Left            =   5760
      Picture         =   "Form_Impressora.frx":030A
      Style           =   1  'Graphical
      TabIndex        =   23
      Top             =   990
      Width           =   1815
   End
   Begin VB.CommandButton Bt_Imprimir 
      Caption         =   "Imprimir"
      Height          =   825
      Left            =   5760
      Picture         =   "Form_Impressora.frx":0614
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   135
      Width           =   1815
   End
   Begin VB.CommandButton Bt_Cancelar 
      Caption         =   "Cancelar"
      Height          =   825
      Left            =   5760
      Picture         =   "Form_Impressora.frx":091E
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   2250
      Width           =   1815
   End
   Begin VB.Frame Frame4 
      Caption         =   "Cópias:"
      Height          =   1005
      Left            =   3645
      TabIndex        =   18
      Top             =   1035
      Width           =   1950
      Begin VB.TextBox Txt_Copias 
         Height          =   285
         Left            =   540
         MaxLength       =   10
         TabIndex        =   19
         Top             =   540
         Width           =   1140
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Num. cópias:"
         Height          =   195
         Left            =   90
         TabIndex        =   20
         Top             =   270
         Width           =   960
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "Orientação:"
      Height          =   1005
      Left            =   45
      TabIndex        =   9
      Top             =   1035
      Width           =   3480
      Begin VB.ComboBox Cb_Orientacao 
         Height          =   315
         ItemData        =   "Form_Impressora.frx":0C28
         Left            =   945
         List            =   "Form_Impressora.frx":0C32
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   585
         Width           =   1950
      End
      Begin VB.ComboBox Cb_Qualidade 
         Height          =   315
         ItemData        =   "Form_Impressora.frx":0C49
         Left            =   945
         List            =   "Form_Impressora.frx":0C5C
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   180
         Width           =   2445
      End
      Begin VB.Label Label4 
         Alignment       =   1  'Right Justify
         Caption         =   "Orientação:"
         Height          =   195
         Left            =   45
         TabIndex        =   13
         Top             =   630
         Width           =   825
      End
      Begin VB.Label Lb_Qualidade 
         Alignment       =   1  'Right Justify
         Caption         =   "Qualidade:"
         Height          =   195
         Left            =   90
         TabIndex        =   12
         Top             =   225
         Width           =   780
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Fonte:"
      Height          =   1005
      Left            =   45
      TabIndex        =   2
      Top             =   2070
      Width           =   4200
      Begin VB.ComboBox Cb_Tamanho 
         Height          =   315
         ItemData        =   "Form_Impressora.frx":0CB5
         Left            =   855
         List            =   "Form_Impressora.frx":0D04
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   585
         Width           =   825
      End
      Begin VB.ComboBox Cb_Fonte 
         Height          =   315
         Left            =   855
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   180
         Width           =   1680
      End
      Begin VB.CheckBox Chk_Sublinhado 
         Caption         =   "Sublinhado"
         Height          =   195
         Left            =   2790
         TabIndex        =   5
         Top             =   720
         Width           =   1095
      End
      Begin VB.CheckBox Chk_Italico 
         Caption         =   "Italico"
         Height          =   195
         Left            =   2790
         TabIndex        =   4
         Top             =   450
         Width           =   825
      End
      Begin VB.CheckBox Chk_Negrito 
         Caption         =   "Negrito"
         Height          =   240
         Left            =   2790
         TabIndex        =   3
         Top             =   180
         Width           =   870
      End
      Begin VB.Label Label3 
         Alignment       =   1  'Right Justify
         Caption         =   "Tamanho:"
         Height          =   195
         Left            =   45
         TabIndex        =   17
         Top             =   675
         Width           =   735
      End
      Begin VB.Label Label2 
         Alignment       =   1  'Right Justify
         Caption         =   "Fonte:"
         Height          =   195
         Left            =   225
         TabIndex        =   16
         Top             =   270
         Width           =   510
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Impressora:"
      Height          =   960
      Left            =   45
      TabIndex        =   0
      Top             =   45
      Width           =   5550
      Begin VB.Label Lb_Porta 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Nenhuma>"
         Height          =   285
         Left            =   990
         TabIndex        =   8
         Top             =   540
         Width           =   4290
      End
      Begin VB.Label Lb_Impressora 
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "<Nenhuma>"
         Height          =   285
         Left            =   990
         TabIndex        =   7
         Top             =   180
         Width           =   4290
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Porta:"
         Height          =   195
         Left            =   405
         TabIndex        =   6
         Top             =   540
         Width           =   510
      End
      Begin VB.Label Label1 
         Alignment       =   1  'Right Justify
         Caption         =   "Impressora:"
         Height          =   240
         Left            =   90
         TabIndex        =   1
         Top             =   225
         Width           =   825
      End
   End
End
Attribute VB_Name = "Form_Impressora"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Status As Boolean
Dim tmp As Integer
Dim TmpFont As String

Private Sub Bt_Cancelar_Click()
    Unload Me
    Status = False
End Sub


Private Sub Bt_Imprimir_Click()
On Error GoTo TrtErroI
    Set ObjPreview = Printer
    ObjPreview.PrintQuality = CI.Qualidade
    ObjPreview.Orientation = CI.Orientacao
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontSize = CI.tFonte
    If Trim(Txt_Copias.Text) <> "" Then
        ObjPreview.Copies = Trim(Txt_Copias.Text)
    End If
    CI.Negrito = IIf(Chk_Negrito.Value = 0, False, True)
    CI.Italico = IIf(Chk_Italico.Value = 0, False, True)
    CI.Italico = IIf(Chk_Sublinhado.Value = 0, False, True)
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Italico
    Status = True
    CI.Preview = False
    Unload Form_Impressora
    'Set ObjPreview = Nothing
    Exit Sub
TrtErroI:
    MsgBox Err.Description, vbInformation, Err.Number
    Resume Next
End Sub

Private Sub Bt_Visualizar_Click()
On Error GoTo TrtErroV
    Set ObjPreview = Form_Preview.Pb_Folha
    'ObjPreview.PrintQuality = Qualidade
    Printer.Orientation = CI.Orientacao
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontSize = CI.tFonte
    'ObjPreview.Copies = IIf(Txt_Copias.Text = "", 1, Txt_Copias.Text)
    CI.Negrito = IIf(Chk_Negrito.Value = 0, False, True)
    CI.Italico = IIf(Chk_Italico.Value = 0, False, True)
    CI.Sublinhado = IIf(Chk_Sublinhado.Value = 0, False, True)
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Italico
    Status = True
    CI.Preview = True
    'Unload Me
    Unload Form_Impressora
    'Set ObjPreview = Nothing
        Exit Sub
TrtErroV:
    Select Case Err.Number
        Case 396
            Resume Next
        Case Else
            MsgBox Err.Description, vbInformation, Err.Number
    End Select
    
End Sub
Private Sub Cb_Fonte_Click()
    CI.Fonte = Cb_Fonte.Text
End Sub

Private Sub Cb_Orientacao_Click()
    CI.Orientacao = IIf(Cb_Orientacao.Text = "Retrato", 1, 2)
End Sub

Private Sub Cb_Qualidade_Click()
    Select Case Cb_Qualidade.Text
        Case "Resolução Rascunho"
            CI.Qualidade = -1
        Case "Resolução Baixa"
            CI.Qualidade = -2
        Case "Resolução Media"
            CI.Qualidade = -3
        Case "Resolução Alta"
            CI.Qualidade = -4
        Case "Personalizada"
            CI.Qualidade = Printer.PrintQuality
        Case Else
            MsgBox "Erro ao localizar a Qualidade", vbInformation, "CESNet - Atenção"
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Cb_Tamanho_Click()
    CI.tFonte = Cb_Tamanho.Text
End Sub

Private Sub Cb_Tamanho_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
Private Sub Form_Load()
    On Error GoTo ErroTrat
    If Trim(Printer.DeviceName) = Empty Then
        MsgBox "Não existe nenhuma impressora instalada. Por favor instale uma antes de continuar...", vbInformation, "CESNet - Aviso"
        Unload Me
        Exit Sub
    End If
    Lb_Impressora.Caption = Printer.DeviceName
    Lb_Porta.Caption = Printer.Port
    For tmp = 1 To Printer.FontCount
        Cb_Fonte.AddItem (Printer.Fonts(tmp))
        If LCase(Printer.Fonts(tmp)) = "arial" Then
            TmpFont = "Arial"
        End If
    Next
        
    Cb_Fonte.Text = IIf(Trim(TmpFont) = "", Printer.Font, TmpFont)
    
    Cb_Tamanho.Text = IIf(Int(Printer.FontSize) < 10, 10, Int(Printer.FontSize))
    Cb_Orientacao = IIf(Printer.Orientation = 1, "Retrato", "Paisagem")
    Select Case Printer.PrintQuality
        Case -1
            Cb_Qualidade.Text = "Resolução Rascunho"
        Case -2
            Cb_Qualidade.Text = "Resolução Baixa"
        Case -3
            Cb_Qualidade.Text = "Resolução Media"
        Case -4
            Cb_Qualidade.Text = "Resolução Alta"
        Case Else
            Cb_Qualidade.Text = "Personalizada"
    End Select
    Chk_Negrito.Value = IIf(Printer.FontBold = True, 1, 0)
    Chk_Italico.Value = IIf(Printer.FontItalic = True, 1, 0)
    Chk_Sublinhado.Value = IIf(Printer.FontUnderline = True, 1, 0)
    Exit Sub
ErroTrat:
    Call RegLogErros(Err.Number, Err.Description, "Form_Impressora", UsuarioID)
    If Err.Number = 482 Or Err.Number = 484 Then
        MsgBox "Nenhuma Impressora instalada por favor verifique.", vbInformation, "Aviso"
        Exit Sub
    End If

End Sub


Private Sub Txt_Copias_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If KeyAscii < 48 Or KeyAscii > 57 Then
        KeyAscii = 0
    End If
End Sub
Public Function LoadFormCI(Optional btImp As Boolean, Optional btPreview As Boolean, _
                      Optional Qualidade As Boolean, Optional Orientacao As Boolean, _
                      Optional Copias As Boolean, Optional Fonte As Boolean, _
                      Optional TamFonte As Boolean, Optional Negrito As Boolean, _
                      Optional Italico As Boolean, Optional Sublinhado As Boolean) As Boolean
    Bt_Imprimir.Enabled = btImp
    Bt_Visualizar.Enabled = btPreview
    Cb_Qualidade.Enabled = Qualidade
    Cb_Orientacao.Enabled = Orientacao
    Txt_Copias.Enabled = Copias
    Cb_Fonte.Enabled = Fonte
    Cb_Tamanho.Enabled = TamFonte
    Chk_Negrito.Enabled = Negrito
    Chk_Italico.Enabled = Italico
    Chk_Sublinhado.Enabled = Sublinhado
    Form_Impressora.Show 1
    ''Debug.Print ObjPreview
    If CI.Preview = True And Status = True Then
        Form_Preview.Show 1
    End If
    
    LoadFormCI = Status
End Function

