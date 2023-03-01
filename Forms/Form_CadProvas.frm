VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form_CadProvas 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "CESNet - Cadastro de Provas"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9315
   Icon            =   "Form_CadProvas.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   6330
   ScaleWidth      =   9315
   Begin VB.ComboBox Cb_Modulo 
      Height          =   315
      Left            =   5370
      TabIndex        =   8
      Top             =   4230
      Width           =   3555
   End
   Begin VB.ComboBox Cb_Serie 
      Height          =   315
      Left            =   720
      TabIndex        =   7
      Top             =   4260
      Width           =   3555
   End
   Begin VB.ComboBox Cb_Disciplina 
      Height          =   315
      Left            =   5400
      TabIndex        =   6
      Top             =   3780
      Width           =   3555
   End
   Begin VB.ComboBox Cb_Ensino 
      Height          =   315
      Left            =   720
      TabIndex        =   5
      Top             =   3780
      Width           =   3555
   End
   Begin VB.Frame Frame2 
      Caption         =   "Prova:"
      Height          =   1455
      Left            =   60
      TabIndex        =   4
      Top             =   4740
      Width           =   8955
      Begin VB.TextBox Txt_Pag 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   20
         TabIndex        =   14
         Top             =   1020
         Width           =   2295
      End
      Begin VB.TextBox Txt_Assunto 
         Enabled         =   0   'False
         Height          =   285
         Left            =   900
         MaxLength       =   100
         TabIndex        =   13
         Top             =   600
         Width           =   7875
      End
      Begin MSMask.MaskEdBox Meb_NProva 
         Height          =   255
         Left            =   900
         TabIndex        =   12
         Top             =   240
         Width           =   555
         _ExtentX        =   979
         _ExtentY        =   450
         _Version        =   393216
         PromptInclude   =   0   'False
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   3
         Mask            =   "###"
         PromptChar      =   "_"
      End
      Begin VB.Label Label7 
         Alignment       =   1  'Right Justify
         Caption         =   "Página(s):"
         Height          =   195
         Left            =   60
         TabIndex        =   11
         Top             =   1080
         Width           =   735
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Caption         =   "Assunto:"
         Height          =   195
         Left            =   120
         TabIndex        =   10
         Top             =   660
         Width           =   675
      End
      Begin VB.Label Label5 
         Alignment       =   1  'Right Justify
         Caption         =   "Nº Prova:"
         Height          =   195
         Left            =   60
         TabIndex        =   9
         Top             =   300
         Width           =   735
      End
   End
   Begin MSComctlLib.Toolbar Tb_Menu 
      Height          =   600
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   10035
      _ExtentX        =   17701
      _ExtentY        =   1058
      ButtonWidth     =   1296
      ButtonHeight    =   953
      AllowCustomize  =   0   'False
      Appearance      =   1
      ImageList       =   "IL_Menu"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   7
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
            Caption         =   "Imprimir"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Gravar"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Cancelar"
            ImageIndex      =   2
         EndProperty
      EndProperty
      Begin MSComctlLib.ImageList IL_Menu 
         Left            =   6255
         Top             =   45
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   32
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":030A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":0624
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":093E
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":0C58
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":0F72
               Key             =   ""
            EndProperty
            BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":128C
               Key             =   ""
            EndProperty
            BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":15A6
               Key             =   ""
            EndProperty
            BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":18C0
               Key             =   ""
            EndProperty
            BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":1BDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":1EF4
               Key             =   ""
            EndProperty
            BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":220E
               Key             =   ""
            EndProperty
            BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":2528
               Key             =   ""
            EndProperty
            BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":2842
               Key             =   ""
            EndProperty
            BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":2B5C
               Key             =   ""
            EndProperty
            BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":2E76
               Key             =   ""
            EndProperty
            BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":3190
               Key             =   ""
            EndProperty
            BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":34AA
               Key             =   ""
            EndProperty
            BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":37C4
               Key             =   ""
            EndProperty
            BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":3ADE
               Key             =   ""
            EndProperty
            BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":3DF8
               Key             =   ""
            EndProperty
            BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":4112
               Key             =   ""
            EndProperty
            BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":442C
               Key             =   ""
            EndProperty
            BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":4746
               Key             =   ""
            EndProperty
            BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":4A60
               Key             =   ""
            EndProperty
            BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":4D7A
               Key             =   ""
            EndProperty
            BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":5094
               Key             =   ""
            EndProperty
            BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":53AE
               Key             =   ""
            EndProperty
            BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":56C8
               Key             =   ""
            EndProperty
            BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":59E2
               Key             =   ""
            EndProperty
            BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":5CFC
               Key             =   ""
            EndProperty
            BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":6016
               Key             =   ""
            EndProperty
            BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Form_CadProvas.frx":6330
               Key             =   ""
            EndProperty
         EndProperty
      End
   End
   Begin MSFlexGridLib.MSFlexGrid MSFG_Provas 
      Height          =   2595
      Left            =   60
      TabIndex        =   17
      Top             =   1020
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   4577
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      SelectionMode   =   1
      AllowUserResizing=   2
      FormatString    =   $"Form_CadProvas.frx":664A
   End
   Begin VB.Label Label8 
      Alignment       =   2  'Center
      BackColor       =   &H00C00000&
      Caption         =   "CADASTRO DE PROVAS"
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
      TabIndex        =   15
      Top             =   600
      Width           =   9315
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      Caption         =   "Modulo:"
      Height          =   195
      Left            =   4680
      TabIndex        =   3
      Top             =   4320
      Width           =   675
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      Caption         =   "Série:"
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   4320
      Width           =   435
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      Caption         =   "Disciplina:"
      Height          =   195
      Left            =   4650
      TabIndex        =   1
      Top             =   3870
      Width           =   705
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      Caption         =   "Curso:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   3840
      Width           =   555
   End
End
Attribute VB_Name = "Form_CadProvas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim RsTrafego As Recordset
Dim RsProvas As Recordset

Dim RsEnsino As Recordset
Dim RsDisciplina As Recordset
Dim RsSerie As Recordset
Dim RsModulo As Recordset

Dim RsTMP As Recordset
Dim tmp As String

Dim Acao As String
Dim lin As String

Dim Ref, EnsinoID, DisciplinaID, SerieID, ModuloID As String






Private Sub Tb_Menu_ButtonClick(ByVal Button As MSComctlLib.Button)
    
    Select Case Button.Index
        Case 1 'Novo
            Acao = 1
            If ChkAcesso(Me.Name, "N") = False Then Exit Sub
            HDFormulario (True)
            hdMenu (False)
            LimpForm
            'MSFG_Provas.Enabled = False

        Case 2 'Alterar
            Acao = 2
            If Meb_NProva.Text = "" Then
                Exit Sub
            End If

            If ChkAcesso(Me.Name, "A") = False Then Exit Sub
            'If Trim(txtDescr.Text) = "" Then Exit Sub
            HDFormulario (True)
            hdMenu (False)
    
            'MSFG_Provas.Enabled = False

            
        Case 3 'Excluir
            Acao = 3
            If ChkAcesso(Me.Name, "E") = False Then Exit Sub
            If Meb_NProva.Text = "" Then
                Exit Sub
            End If

            ExcProva
            
            'LimpForm
            'ExcluirCampo
            
        Case 4 'Imprimir
            Acao = 4
            If ChkAcesso(Me.Name, "I") = False Then Exit Sub
            'ImprimirListagem
        Case 6 'Gravar
            'GravarDados
            If ValidarSoftware("Provas") = False Then Exit Sub
            
            If Trim(Meb_NProva.Text) = "" Then
                MsgBox "O campo NUMERO DA PROVA não pode ser deixado em branco. Por favor, verifique.", vbInformation, "CESNet - Aviso!"
                Exit Sub
            End If
            GrvProva
            LimpForm
            HDFormulario (False)
            hdMenu (True)
            'MstDados
        Case 7 'Cancelar
            Acao = 0
            HDFormulario (False)
            hdMenu (True)
            LimpForm
    End Select
End Sub



    
    

Private Sub ExcProva()
    tmp = MsgBox("Deseja remover esta prova?" & Chr(13) & Chr(13) & "- Prova nº.: " & Meb_NProva.Text & Chr(13) & "- Assunto: " & Txt_Assunto.Text, vbYesNo, "CESNet - Aviso")
        If tmp = 6 Then
            PegarReferencias
            Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID & " AND ModuloID = " & ModuloID)
            RsTrafego.MoveFirst
            Ref = RsTrafego.Fields("RefTrafegoID")
            'pEGA A PROVA SEGUNDO A REF
            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & Ref)
            If RsProvas.BOF And RsProvas.EOF Then
                    MsgBox "Erro na busca da(s) prova(s) com referencia: " & Ref, vbInformation, "CESNet - Aviso!"
                    Exit Sub
                Else
                    'busca o num da prova
                    RsProvas.FindFirst "Nprova = '" & Meb_NProva.Text & "'"
                    If RsProvas.NoMatch Then
                            MsgBox "Erro ao procurar a prova: " & Meb_NProva.Text, vbInformation, "CESNet - Aviso!"
                            Exit Sub
                        Else
                            'apaga a prova
                            RsProvas.Delete
                            MsgBox "Prova apagada com sucesso!", vbInformation, "CESNet"
                            'Checa se existe mais refencias na tab provas caso nao tenha
                            'apaga no trafego o modulo
                            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID) 'RefTrafegoID = " & Ref)
                            If RsProvas.BOF And RsProvas.EOF Then
                                RsTrafego.Delete
                                Set RsTMP = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
                                If RsTMP.BOF And RsTMP.EOF Then
                                    BD.Execute "DELETE * FROM GradeEnsinoSeries WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID
                                    Set RsTMP = BD.OpenRecordset("SELECT * FROM GradeEnsinoSeries WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID)
                                    If RsTMP.BOF And RsTMP.EOF Then
                                        BD.Execute "DELETE * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID
                                    End If
                                End If
                            End If
                    End If
            End If
            
            
            'RsProvas.FindFirst "RefTrafegoID ='" & Ref & "' AND Nprova = '" & Meb_NProva.Text & "'"
            'If RsProvas.NoMatch Then
            '        Exit Sub
            '    Else
            '        RsProvas.Delete
            'End If
        
        End If
    LstProvas
End Sub

Private Sub GrvProva()
    
    Select Case Acao
        Case 1 'INCLUSAO
            If PegarReferencias = False Then Exit Sub
            
            Set RsTMP = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID= " & EnsinoID & " AND DisciplinaId = " & DisciplinaID)
            If RsTMP.BOF And RsTMP.EOF Then
                RsTMP.AddNew
                RsTMP.Fields("EnsinoID") = EnsinoID
                RsTMP.Fields("DisciplinaID") = DisciplinaID
                RsTMP.Fields("Obrigatoria") = True
                RsTMP.Update
                RsTMP.Close
            End If
            Set RsTMP = BD.OpenRecordset("SELECT * FROM GradeEnsinoSeries WHERE EnsinoID = " & EnsinoID & " AND DisciplinaId = " & DisciplinaID & "  AND SerieID = " & SerieID)
            If RsTMP.BOF And RsTMP.EOF Then
                RsTMP.AddNew
                RsTMP.Fields("EnsinoID") = EnsinoID
                RsTMP.Fields("DisciplinaID") = DisciplinaID
                RsTMP.Fields("SerieID") = SerieID
                RsTMP.Update
                RsTMP.Close
            End If
            Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaId = " & DisciplinaID & " AND SerieID = " & SerieID & "  AND ModuloID = " & ModuloID)
            If RsTrafego.BOF And RsTrafego.EOF Then
                    With RsTrafego
                        .AddNew
                        .Fields("EnsinoID") = EnsinoID
                        .Fields("DisciplinaID") = DisciplinaID
                        .Fields("SerieID") = SerieID
                        .Fields("ModuloID") = ModuloID
                        .Update
                        .FindFirst "EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID & " AND ModuloID = " & ModuloID
                        If .NoMatch Then
                            Else
                                .MoveFirst
                                Ref = .Fields("RefTrafegoID")
                        End If
                    End With
                Else
                    RsTrafego.MoveFirst
                    Ref = RsTrafego.Fields("RefTrafegoID")
                    Set RsTMP = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & Ref)
                    RsTMP.FindFirst "NProva = '" & Meb_NProva.Text & "'"
                    If RsTMP.NoMatch Then
                            RsTMP.Close
                        Else
                            MsgBox "Este Numero de Prova ja foi cadastrada. Por favor verifique!", vbInformation, "CESNet - Aviso!"
                            Meb_NProva.SetFocus
                            RsTMP.Close
                            Exit Sub
                    End If
                    
            End If
            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas")  ' WHERE RefTrafegoID ='" & Ref & "'")
            With RsProvas
                .AddNew
                .Fields("RefTrafegoID") = Ref
                .Fields("EnsinoID") = EnsinoID
                .Fields("DisciplinaID") = DisciplinaID
                .Fields("SerieID") = SerieID
                .Fields("ModuloID") = ModuloID
                .Fields("NProva") = Meb_NProva.Text
                .Fields("Assunto") = IIf(Txt_Assunto.Text = "", " ", Txt_Assunto.Text)
                .Fields("Pag") = IIf(Txt_Pag.Text = "", " ", Txt_Pag.Text)
                .Update
            End With
            LstProvas
            Acao = 0
        
        Case 2 ' ALTERACAO
            PegarReferencias
            'Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaId = " & DisciplinaID & " AND SerieID = " & SerieID & " AND ModuloID = " & ModuloID)
            'If RsTrafego.BOF And RsTrafego.EOF Then
            '        MsgBox "Erro no acesso ao sistema tente novamente", vbInformation, "CESNet - Aviso!"
            '    Else
            '        RsTrafego.MoveFirst
            '        Ref = RsTrafego.Fields("RefTrafegoID")
            'End If
            
            'Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & Ref & " AND NProva = '" & MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0) & "' AND Assunto= '" & MSFG_Provas.TextMatrix(MSFG_Provas.Row, 1) & "'")
            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE " & _
                                            "EnsinoID = " & EnsinoID & _
                                            " And DisciplinaID = " & DisciplinaID & _
                                            " And SerieID = " & SerieID & _
                                            " And ModuloID = " & ModuloID & _
                                            " AND NProva = '" & MSFG_Provas.TextMatrix(MSFG_Provas.Row, 0) & _
                                            "' AND Assunto= '" & MSFG_Provas.TextMatrix(MSFG_Provas.Row, 1) & "'")
            '
            If RsProvas.BOF And RsProvas.EOF Then
                    MsgBox "Erro ao localizar prova.", vbInformation, "CESNet"
                    Exit Sub
                Else
                    RsProvas.MoveFirst
                    'MsgBox "Este Numero de Prova já foi cadastrada. Por favor verifique!", vbInformation, "CESNet - Aviso!"
                    'Meb_NProva.SetFocus
                    'RsProvas.Close
                    'Exit Sub
            End If
            
            
            With RsProvas
                .Edit
                '.Fields("RefTrafegoID") = Ref
                .Fields("EnsinoID") = EnsinoID
                .Fields("DisciplinaID") = DisciplinaID
                .Fields("SerieID") = SerieID
                .Fields("ModuloID") = ModuloID
                .Fields("NProva") = Meb_NProva.Text
                .Fields("Assunto") = IIf(Txt_Assunto.Text = "", " ", Txt_Assunto.Text)
                .Fields("Pag") = IIf(Txt_Pag.Text = "", " ", Txt_Pag.Text)
                .Update
            End With
            LstProvas
            Acao = 0
    End Select
End Sub

Private Sub Cb_Ensino_Click()
    LstProvas
End Sub

Private Sub Cb_Ensino_DropDown()
    Cb_Ensino.Clear
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino ORDER BY Descr")
    If RsEnsino.BOF And RsEnsino.BOF Then
            MsgBox "Não existe nenhum Ensino cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsEnsino.MoveFirst
            Do Until RsEnsino.EOF
                Cb_Ensino.AddItem (RsEnsino.Fields("Descr"))
                RsEnsino.MoveNext
            Loop
    End If

End Sub




Private Sub Cb_Ensino_KeyPress(KeyAscii As Integer)
 
    If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub Cb_Disciplina_Click()
    LstProvas
End Sub

Private Sub Cb_Disciplina_DropDown()
    Cb_Disciplina.Clear
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina ORDER BY Descr")
    If RsDisciplina.BOF And RsDisciplina.BOF Then
            MsgBox "Não existe nenhum Disciplina cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsDisciplina.MoveFirst
            Do Until RsDisciplina.EOF
                Cb_Disciplina.AddItem (RsDisciplina.Fields("Descr"))
                RsDisciplina.MoveNext
            Loop
    End If
End Sub

Private Sub Cb_Disciplina_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub Cb_Modulo_Click()
    LstProvas
End Sub

Private Sub Cb_Modulo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub Cb_Serie_Click()
    LstProvas
End Sub

Private Sub Cb_Serie_DropDown()
    Cb_Serie.Clear
    Set RsSerie = BD.OpenRecordset("SELECT * FROM Serie ORDER BY Descr")
    If RsSerie.BOF And RsSerie.BOF Then
            MsgBox "Não existe nenhum Serie cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsSerie.MoveFirst
            Do Until RsSerie.EOF
                Cb_Serie.AddItem (RsSerie.Fields("Descr"))
                RsSerie.MoveNext
            Loop
    End If
End Sub

Private Sub Cb_Modulo_DropDown()
    Cb_Modulo.Clear
    Set RsModulo = BD.OpenRecordset("SELECT * FROM Modulo ORDER BY ID")
    If RsModulo.BOF And RsModulo.BOF Then
            MsgBox "Não existe nenhum Modulo cadastrado. Pro favor cadastre antes de incluir provas.", vbInformation, "CESNet - Aviso!"
            Exit Sub
        Else
            RsModulo.MoveFirst
            Do Until RsModulo.EOF
                Cb_Modulo.AddItem (RsModulo.Fields("Descr"))
                RsModulo.MoveNext
            Loop
    End If
End Sub


Private Sub Cb_Serie_KeyPress(KeyAscii As Integer)

    If KeyAscii = 13 Then
            SendKeys "{TAB}"
            KeyAscii = 0
        Else
            KeyAscii = 0
    End If
End Sub

Private Sub Form_Activate()
    If ChkAcesso(Me.Name, "C") = False Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    Me.top = 0
    Me.left = 0
    Acao = 0
    HDFormulario (False)
    hdMenu (True)

End Sub
Private Sub hdMenu(op As Boolean)

    Tb_Menu.Buttons(1).Enabled = op
    Tb_Menu.Buttons(2).Enabled = op
    Tb_Menu.Buttons(3).Enabled = op
    Tb_Menu.Buttons(4).Enabled = False 'op
    
    Tb_Menu.Buttons(6).Enabled = IIf(op = False, True, False)
    Tb_Menu.Buttons(7).Enabled = IIf(op = False, True, False)
    
End Sub

Private Sub Meb_NProva_GotFocus()
    Meb_NProva.SelStart = 0
    Meb_NProva.SelLength = 3
End Sub

Private Sub Meb_NProva_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Meb_NProva_LostFocus()
    If Meb_NProva.Text = "" Then
        Else
            Meb_NProva.Text = Mid(String(3, "0"), 1, 3 - Len(Meb_NProva.Text)) & Meb_NProva.Text
    End If
End Sub



Private Sub MSFG_Provas_Click()
    With MSFG_Provas
        Meb_NProva.Text = .TextMatrix(.Row, 0)
        Txt_Assunto.Text = .TextMatrix(.Row, 1)
        Txt_Pag.Text = .TextMatrix(.Row, 2)
    End With
End Sub

Private Sub Txt_Assunto_GotFocus()
    Txt_Assunto.SelStart = 0
    Txt_Assunto.SelLength = Len(Txt_Assunto.Text)
End Sub


Private Sub Txt_Assunto_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub

Private Sub Txt_Pag_GotFocus()
    Txt_Pag.SelStart = 0
    Txt_Pag.SelLength = Len(Txt_Pag.Text)
End Sub

Private Sub Txt_Pag_KeyPress(KeyAscii As Integer)
    KeyAscii = Asc(UCase(Chr(KeyAscii)))
    If KeyAscii = 13 Then
        SendKeys "{TAB}"
        KeyAscii = 0
    End If
End Sub
Private Sub LimpForm()
    'MSFG_Provas.Rows = 1
    'MSFG_Provas.Rows = 2
    'Cb_Ensino.Clear
    'Cb_Disciplina.Clear
    'Cb_Serie.Clear
    'Cb_Modulo.Clear
    
    Meb_NProva.Text = ""
    Txt_Assunto.Text = ""
    Txt_Pag.Text = ""
End Sub

Private Sub DesCbs()
    Cb_Ensino.Enabled = False
    Cb_Disciplina.Enabled = False
    Cb_Serie.Enabled = False
    Cb_Modulo.Enabled = False
End Sub
Private Sub HabCbs()
    Cb_Ensino.Enabled = True
    Cb_Disciplina.Enabled = True
    Cb_Serie.Enabled = True
    Cb_Modulo.Enabled = True
End Sub

Private Sub HDFormulario(op As Boolean)
    MSFG_Provas.Enabled = IIf(op = True, False, True)
    
    'HDFormulario op
    If Acao = 1 Then
            Cb_Ensino.Enabled = op
            Cb_Disciplina.Enabled = op
            Cb_Serie.Enabled = op
            Cb_Modulo.Enabled = op
        Else
            Cb_Ensino.Enabled = IIf(op = True, False, True)
            Cb_Disciplina.Enabled = IIf(op = True, False, True)
            Cb_Serie.Enabled = IIf(op = True, False, True)
            Cb_Modulo.Enabled = IIf(op = True, False, True)
    End If
    Meb_NProva.Enabled = op
    Txt_Assunto.Enabled = op
    Txt_Pag.Enabled = op
End Sub

Private Sub LstProvas()
    MSFG_Provas.Rows = 1
    MSFG_Provas.Rows = 2
    If Cb_Ensino.Text = "" Or Cb_Disciplina.Text = "" Or Cb_Serie.Text = "" Or Cb_Modulo.Text = "" Then
        Exit Sub
    End If
    PegarReferencias
    'Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID & " AND ModuloID = " & ModuloID)
    'If RsTrafego.BOF And RsTrafego.EOF Then
    '        MSFG_Provas.Rows = 1
    '        MSFG_Provas.Rows = 2
    '        Exit Sub
    '    Else
    '        MSFG_Provas.Rows = 1
    '        MSFG_Provas.Rows = 2
    '        tmp = RsTrafego.Fields("RefTrafegoID")
            Set RsProvas = BD.OpenRecordset("SELECT * FROM Provas WHERE " & _
            "EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID & " AND ModuloID = " & ModuloID & _
            " ORDER BY NProva")
            If RsProvas.BOF And RsProvas.EOF Then
                    Exit Sub
                Else
                    lin = 1
                    RsProvas.MoveFirst
                    Do Until RsProvas.EOF
                        MSFG_Provas.TextMatrix(lin, 0) = RsProvas.Fields("NProva")
                        MSFG_Provas.TextMatrix(lin, 1) = RsProvas.Fields("Assunto")
                        MSFG_Provas.TextMatrix(lin, 2) = RsProvas.Fields("Pag")
                        MSFG_Provas.Rows = MSFG_Provas.Rows + 1
                        lin = lin + 1
                        RsProvas.MoveNext
                    Loop
                    MSFG_Provas.Rows = MSFG_Provas.Rows - 1
            End If
    'End If
End Sub
Private Function PegarReferencias() As Boolean
    If Cb_Ensino.Text = "" Or Cb_Disciplina.Text = "" Or Cb_Serie.Text = "" Or Cb_Modulo.Text = "" Then
        MsgBox "Favor verificar um dos item ENSINO, SERIE, DISCIPLINA OU MODULO, pois um deles está vazio...", vbInformation, "Atenção"
        PegarReferencias = False
        Exit Function
    End If
    RsEnsino.FindFirst "Descr ='" & Cb_Ensino.Text & "'"
    EnsinoID = RsEnsino.Fields("ID")
    
    RsDisciplina.FindFirst "Descr ='" & Cb_Disciplina.Text & "'"
    DisciplinaID = RsDisciplina.Fields("ID")
    
    RsSerie.FindFirst "Descr ='" & Cb_Serie.Text & "'"
    SerieID = RsSerie.Fields("ID")
    
    RsModulo.FindFirst "Descr ='" & Cb_Modulo.Text & "'"
    ModuloID = RsModulo.Fields("ID")
    PegarReferencias = True
End Function


