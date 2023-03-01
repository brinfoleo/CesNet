VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form_MatriculaAviso 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - AVISO ou BLOQUEIO de Matricula"
   ClientHeight    =   6690
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9330
   Icon            =   "Form_MatriculaAviso.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6690
   ScaleWidth      =   9330
   Begin MSFlexGridLib.MSFlexGrid msfgAvisos 
      Height          =   1995
      Left            =   60
      TabIndex        =   5
      Top             =   1020
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   3519
      _Version        =   393216
      Cols            =   6
      SelectionMode   =   1
      AllowUserResizing=   1
      FormatString    =   $"Form_MatriculaAviso.frx":030A
   End
   Begin VB.Frame Frame_Avisar 
      Height          =   1215
      Left            =   120
      TabIndex        =   3
      Top             =   5280
      Width           =   9135
      Begin MSComCtl2.DTPicker DTP_Aviso 
         Height          =   315
         Left            =   1980
         TabIndex        =   14
         Top             =   360
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60686337
         CurrentDate     =   40665
      End
      Begin VB.CheckBox Chk_Bloquear 
         Caption         =   "Bloquear a partir de:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   180
         TabIndex        =   13
         Top             =   780
         Width           =   1755
      End
      Begin VB.CheckBox Chk_Avisar 
         Caption         =   "Avisar a partir de:"
         Enabled         =   0   'False
         Height          =   195
         Left            =   180
         TabIndex        =   4
         Top             =   360
         Width           =   1635
      End
      Begin MSComCtl2.DTPicker DTP_Bloqueio 
         Height          =   315
         Left            =   1980
         TabIndex        =   12
         Top             =   720
         Width           =   1335
         _ExtentX        =   2355
         _ExtentY        =   556
         _Version        =   393216
         Format          =   60686337
         CurrentDate     =   38633
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Texto:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1155
      Left            =   120
      TabIndex        =   1
      Top             =   4080
      Width           =   9135
      Begin VB.TextBox Txt_Texto 
         Enabled         =   0   'False
         Height          =   795
         Left            =   180
         MaxLength       =   200
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   2
         Top             =   240
         Width           =   8835
      End
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   6
      Top             =   0
      Width           =   11175
      _ExtentX        =   19711
      _ExtentY        =   1058
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "IL_Menu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   6
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Novo"
            ImageIndex      =   24
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Alterar"
            ImageIndex      =   32
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Excluir"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList IL_Menu 
         Left            =   7320
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   33
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":03AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":06C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":09E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":0CFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":1016
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":1330
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":164A
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":1964
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":1C7E
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":1F98
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":22B2
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":25CC
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":28E6
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":2C00
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":2F1A
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":3234
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":354E
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":3868
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":3B82
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":3E9C
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":41B6
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":44D0
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":47EA
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":4B04
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":4E1E
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":5138
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":5452
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":576C
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":5A86
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":5DA0
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":60BA
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":63D4
               Key             =   ""
            EndProperty
            BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_MatriculaAviso.frx":66EE
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSComCtl2.DTPicker DTP_Dt 
      Height          =   315
      Left            =   900
      TabIndex        =   7
      Top             =   3180
      Width           =   1395
      _ExtentX        =   2461
      _ExtentY        =   556
      _Version        =   393216
      Enabled         =   0   'False
      Format          =   60686337
      CurrentDate     =   38633
   End
   Begin MSMask.MaskEdBox MebMatricula 
      Height          =   315
      Left            =   900
      TabIndex        =   8
      Top             =   3600
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   556
      _Version        =   393216
      Appearance      =   0
      MaxLength       =   11
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Mask            =   "##.###.####"
      PromptChar      =   "_"
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Matricula:"
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   3660
      Width           =   735
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Data:"
      Height          =   195
      Left            =   300
      TabIndex        =   10
      Top             =   3240
      Width           =   495
   End
   Begin VB.Label Lb_Nome 
      Height          =   195
      Left            =   2340
      TabIndex        =   9
      Top             =   3660
      Width           =   6555
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "AVISO / BLOQUEIO DE MATRICULA"
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
      Top             =   600
      Width           =   9285
   End
End
Attribute VB_Name = "Form_MatriculaAviso"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsMatricula     As Recordset
Dim RsAviso         As Recordset

Dim MatrID          As String
Dim RegID           As Long 'As Integer

Private Sub Chk_Avisar_Click()
    If Chk_Avisar.Enabled = False Then Exit Sub
    If Chk_Avisar.Value = 0 Then
            DTP_Aviso.Enabled = False
        Else
            DTP_Aviso.Enabled = True
    End If

End Sub

Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            RegID = 0
            hdMenu (False)
            HDForm (True)
            LimpForm
            
        Case 2 'Alterar
            If RegID = 0 Then
                MsgBox "Favor selecionar um registro.", vbInformation, "Aviso"
                Exit Sub
            End If
             If Trim(msfgAvisos.TextMatrix(msfgAvisos.Row, 2)) <> "" Then
                MsgBox "Aviso não pode ser alterado!", vbInformation, "Aviso"
                Exit Sub
            End If
            hdMenu (False)
            HDForm (True)
            
        Case 3 'Excluir
            If RegID = 0 Then
                MsgBox "Favor selecionar um registro.", vbInformation, "Aviso"
                Exit Sub
            End If
            Excluir
        Case 5 'Gravar
            Gravar
            LimpForm
            hdMenu (True)
            HDForm (False)
        Case 6 'Cancelar
            hdMenu (True)
            HDForm (False)
    End Select
End Sub
Private Sub LimpForm()
    DTP_Dt.Value = Date
    Txt_Texto.Text = ""
    'Lb_Nome.Caption = ""
    Chk_Avisar.Value = 0
    Chk_Bloquear.Value = 0
    DTP_Aviso.Value = Date
    DTP_Bloqueio.Value = Date
    'Bt_Gravar.Enabled = op
    'Bt_Excluir.Enabled = op
End Sub

Private Sub Excluir()
    If Trim(msfgAvisos.TextMatrix(msfgAvisos.Row, 2)) <> "" Then
        MsgBox "Aviso não pode ser excluido!", vbInformation, "Aviso"
        Exit Sub
    End If
    If MsgBox("Deseja realmente excluir este aviso?", vbInformation + vbYesNo, "CESNet - Aviso") = vbYes Then
        BD.Execute "DELETE * FROM MatriculaAviso WHERE ID = " & RegID
        RegID = 0
        LimpForm
        PgAviso
    End If
End Sub


Private Sub Gravar()
    If Trim(MatrID) = "" Then
        MsgBox "Digite uma matricula!", vbInformation, "Aviso"
        Exit Sub
    End If
    If RegID = 0 Then
            Set RsAviso = BD.OpenRecordset("SELECT * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "'")
            RsAviso.AddNew
        Else
            Set RsAviso = BD.OpenRecordset("SELECT * FROM MatriculaAviso WHERE ID = " & RegID)
            If RsAviso.BOF And RsAviso.EOF Then
                    MsgBox "Erro ao localizar registro para alteração! Operação Cancelada!", vbInformation, "Aviso"
                    Exit Sub
                Else
                    RsAviso.MoveFirst
                    RsAviso.Edit
            End If
    End If
    RsAviso.Fields("MatrID") = MatrID
    RsAviso.Fields("DtInclusao") = DTP_Dt.Value
    RsAviso.Fields("Texto") = Txt_Texto.Text
    RsAviso.Fields("Avisar") = IIf(Chk_Avisar.Value = 1, True, False)
    RsAviso.Fields("DtAvisar") = IIf(Chk_Avisar.Value = 1, DTP_Aviso.Value, Null)
    RsAviso.Fields("Bloquear") = IIf(Chk_Bloquear.Value = 1, True, False)
    RsAviso.Fields("DtBloqueio") = IIf(Chk_Bloquear.Value = 1, DTP_Bloqueio.Value, Null)
    RsAviso.Fields("UsuID") = UsuarioID
    RsAviso.Fields("DtHr") = Now
    RsAviso.Update
    MsgBox "Aviso gravado com sucesso.", vbInformation, "CESNet - Aviso"
    LimpForm
    HDForm (False)
    PgAviso
End Sub

Private Sub Chk_Bloquear_Click()
    If Chk_Bloquear.Enabled = False Then Exit Sub
    If Chk_Bloquear.Value = 0 Then
           DTP_Bloqueio.Enabled = False
        Else
            DTP_Bloqueio.Enabled = True
    End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    
    DTP_Dt.Enabled = False
    DTP_Bloqueio.Enabled = False
    DTP_Aviso.Enabled = False
    
    DTP_Dt.Value = Date
    DTP_Aviso.Value = Date
    DTP_Bloqueio.Value = Date
    
    HDForm (False)
    hdMenu (True)
End Sub

Private Sub MebMatricula_GotFocus()
    MebMatricula.SelStart = 0
    MebMatricula.SelLength = 11
End Sub
Private Sub MebMatricula_KeyPress(KeyAscii As Integer)
    DoEvents
    If KeyAscii = 13 Then
        Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID ='" & MebMatricula.Text & "'")
        With RsMatricula
            If .BOF And .EOF Then
                    MsgBox "Matricula não encontrada.", vbInformation, "CESNet - Aviso!"
                    MatrID = ""
                    'HDForm (False)
                    MebMatricula.SetFocus
                    Exit Sub
                Else
                    'HDForm (True)
                    RsMatricula.MoveFirst
                    MatrID = MebMatricula.Text
                    Lb_Nome.Caption = RsMatricula.Fields("Nome")
                    RsMatricula.Close
                    PgAviso
            End If
        End With
    End If
End Sub

Private Sub HDForm(op As Boolean)
    msfgAvisos.Enabled = IIf(op = True, False, True)
    DTP_Dt.Enabled = op
    Txt_Texto.Enabled = op
    
    Chk_Avisar.Enabled = op
    Chk_Bloquear.Enabled = op
    
    
    
    'Bt_Gravar.Enabled = op
    'Bt_Excluir.Enabled = op
End Sub
Private Sub hdMenu(op As Boolean)
    Tb_Menu.Buttons(1).Enabled = op
    Tb_Menu.Buttons(2).Enabled = op
    Tb_Menu.Buttons(3).Enabled = op
    
    Tb_Menu.Buttons(5).Enabled = IIf(op = True, False, True)
    Tb_Menu.Buttons(6).Enabled = IIf(op = True, False, True)
End Sub
Private Sub PgAviso()
    msfgAvisos.Rows = 1
    Set RsAviso = BD.OpenRecordset("SELECT * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "'")
    If RsAviso.BOF And RsAviso.EOF Then
            'LimpForm
        Else
            RsAviso.MoveFirst
            
            With msfgAvisos
                Do Until RsAviso.EOF
                    .Rows = .Rows + 1
                    .TextMatrix(.Rows - 1, 0) = RsAviso.Fields("id")
                    .TextMatrix(.Rows - 1, 1) = RsAviso.Fields("DtInclusao")
                    .TextMatrix(.Rows - 1, 2) = IIf(IsNull(RsAviso.Fields("Codigo")), " ", RsAviso.Fields("Codigo"))
                    .TextMatrix(.Rows - 1, 3) = RsAviso.Fields("texto")
                    .TextMatrix(.Rows - 1, 4) = IIf(RsAviso.Fields("Avisar") = True, IIf(IsNull(RsAviso.Fields("DtAvisar")), Date, RsAviso.Fields("DtAvisar")), " ")
                    .TextMatrix(.Rows - 1, 5) = IIf(IsNull(RsAviso.Fields("DtBloqueio")), " ", RsAviso.Fields("DtBloqueio"))
                    RsAviso.MoveNext
                   
                Loop
            End With
            'Bt_Gravar.Enabled = op
            'Bt_Excluir.Enabled = op
    End If
End Sub

Private Sub msfgAvisos_Click()
    With msfgAvisos
        RegID = .TextMatrix(.Row, 0)
        DTP_Dt.Value = .TextMatrix(.Row, 1)
        Txt_Texto.Text = .TextMatrix(.Row, 3)
    
        Chk_Avisar.Value = IIf(.TextMatrix(.Row, 4) = "SIM", 1, 0)
        Chk_Bloquear.Value = IIf(Trim(.TextMatrix(.Row, 5)) <> "", 1, 0)
                
        DTP_Bloqueio.Value = IIf(Trim(.TextMatrix(.Row, 5)) <> "", .TextMatrix(.Row, 5), Date)
    End With
End Sub


Private Sub Txt_Texto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
End Sub


