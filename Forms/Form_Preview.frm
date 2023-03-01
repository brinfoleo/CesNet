VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form Form_Preview 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Visualizar"
   ClientHeight    =   3825
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   Icon            =   "Form_Preview.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3825
   ScaleWidth      =   5460
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin ComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   5460
      _ExtentX        =   9631
      _ExtentY        =   741
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   2
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin VB.PictureBox Pb_Dim 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   2925
         ScaleHeight     =   225
         ScaleWidth      =   2085
         TabIndex        =   5
         Top             =   105
         Width           =   2115
      End
   End
   Begin VB.PictureBox Pb_Fundo 
      Height          =   3120
      Left            =   90
      ScaleHeight     =   3060
      ScaleWidth      =   3960
      TabIndex        =   0
      Top             =   540
      Width           =   4020
      Begin VB.HScrollBar HScroll 
         Height          =   420
         Left            =   45
         TabIndex        =   3
         Top             =   2700
         Width           =   3165
      End
      Begin VB.VScrollBar VScroll 
         Height          =   3120
         Left            =   3465
         TabIndex        =   2
         Top             =   0
         Width           =   375
      End
      Begin VB.PictureBox Pb_Folha 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   960
         Left            =   225
         ScaleHeight     =   930
         ScaleWidth      =   975
         TabIndex        =   1
         Top             =   360
         Width           =   1005
      End
   End
   Begin ComctlLib.ImageList ImageList1 
      Left            =   4680
      Top             =   540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   2
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Preview.frx":030A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "Form_Preview.frx":0624
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form_Preview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim Bordas As Integer
Private Sub Form_Load()
On Error GoTo ErroTrat
    'Me.Caption = Me.Caption & "Orientação: " & Printer.Orientation
    Pb_Folha.Height = Printer.Height
    Pb_Folha.Width = Printer.Width
    VScroll.Min = 0
    VScroll.Max = 100
    
    HScroll.Min = 0
    HScroll.Max = 100
    Exit Sub
ErroTrat:
    Call RegLogErros(Err.Number, Err.Description, "Form_Preview", UsuarioID)
    If Err.Number = 482 Or Err.Number = 484 Then
        MsgBox "Nenhuma Impressora instalada por favor verifique", vbInformation, "Aviso"
        Exit Sub
    End If
End Sub
Private Sub Form_Resize()
    If Form_Preview.ScaleHeight = 0 Or Form_Preview.ScaleWidth = 0 Then
        Exit Sub
    End If
        
    Bordas = 100
    Pb_Fundo.Top = Toolbar1.Height + Bordas
    Pb_Fundo.Left = Bordas

    Pb_Fundo.Height = Form_Preview.ScaleHeight - Pb_Fundo.Top - Bordas
    Pb_Fundo.Width = Form_Preview.ScaleWidth - (2 * Bordas)
    
    Pb_Folha.Top = 135
    Pb_Folha.Left = 135
        
    VScroll.Width = 275
    HScroll.Height = VScroll.Width
    
    VScroll.Top = 0
    VScroll.Height = Pb_Fundo.ScaleHeight
    VScroll.Left = Pb_Fundo.ScaleWidth - VScroll.Width
    
    
    HScroll.Left = 0 ' - HScroll.Height
    HScroll.Top = Pb_Fundo.ScaleHeight - HScroll.Height
    HScroll.Width = Pb_Fundo.ScaleWidth - VScroll.Width
    
    
    Pb_Folha.ScaleMode = 7
    Pb_Dim.Print "  Larg.: " & Format(Pb_Folha.ScaleWidth, "#.0") & " x Alt.:" & Format(Pb_Folha.ScaleHeight, "#.0") & " cm"
    Pb_Dim.Width = Pb_Dim.TextWidth("  Larg.: " & Format(Pb_Folha.ScaleWidth, "#.0") & " x Alt.:" & Format(Pb_Folha.ScaleHeight, "#.0") & " cm") + Bordas
    Pb_Dim.Left = Form_Preview.ScaleWidth - (Pb_Dim.Width + Bordas)
    Pb_Folha.ScaleMode = 1
    
    
    'Pb_Folha.FontBold = False
    'Pb_Folha.FontItalic = False
    'Pb_Folha.FontUnderline = False
    'Pb_Folha.Font = "Arial"
    'Pb_Folha.FontSize = 6
    'Pb_Folha.CurrentX = Pb_Folha.ScaleHeight - Printer.TextHeight("CESNet - Programa de Gerenciamento Escolar")
    'Pb_Folha.CurrentY = Pb_Folha.ScaleHeight - Printer.TextWidth("CESNet - Programa de Gerenciamento Escolar")
    'Pb_Folha.Print "CESNet - Programa de Gerenciamento Escolar"
    AtivarScroll
    VScroll_Change
    HScroll_Change
End Sub
Private Sub AtivarScroll()
    DoEvents
    If (Pb_Folha.Height + HScroll.Height + 135) < Pb_Fundo.ScaleHeight Then
            VScroll.Enabled = False
        Else
            VScroll.Enabled = True
    End If
    If (Pb_Folha.Width + VScroll.Width + 135) < Pb_Fundo.ScaleWidth Then
            HScroll.Enabled = False
        Else
            HScroll.Enabled = True
    End If
End Sub





Private Sub Pb_Folha_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Pb_Folha.MousePointer = 12
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As ComctlLib.Button)
    On Error GoTo TratErro
    Select Case Button.Index
        Case 1
            'A propriedade AutroDraw deve estar como true p/ que possa imprimir a imagem
            Printer.PaintPicture Pb_Folha.Image, 0, 0, Pb_Folha.Width, Pb_Folha.Height
            Printer.EndDoc
        Case 2
            Unload Me
        
        'Case 3
        '    ObjPreview.PaintPicture ImageList1.ListImages.Item(1).Picture, Pb_Folha.Width / 2, 1500, 1600, 1600
        '
        '    ObjPreview.Font = "Arial"
        '    ObjPreview.CurrentX = 0
        '    ObjPreview.CurrentY = 0
        '    ObjPreview.FontSize = 12
        '    ObjPreview.Print "Texto Teste"
        '
        '    ObjPreview.CurrentX = 250
        '    ObjPreview.CurrentY = 250
        '   ObjPreview.FontSize = 20
        '   ObjPreview.Print "Texto Teste"

    End Select
    Exit Sub
TratErro:
    MsgBox Err.Description, vbCritical, Err.Number
    Call RegLogErros(Err.Number, Err.Description, Me.Caption, UsuarioID)
    
End Sub

Private Sub VScroll_Change()
    X = Val(Pb_Folha.Height) - (Val(Pb_Fundo.Height) - HScroll.Height - Bordas - 270)
    DoEvents
    Pb_Folha.Top = ((X * VScroll.Value / 100) - 135) * -1 ' ((VScroll.Value * 135) / 100) * -1
End Sub
Private Sub HScroll_Change()
    X = Val(Pb_Folha.Width) - (Val(Pb_Fundo.Width) - VScroll.Width - Bordas - 270)
    DoEvents
    Pb_Folha.Left = ((X * HScroll.Value / 100) - 135) * -1  ' ((VScroll.Value * 135) / 100) * -1
End Sub
