Attribute VB_Name = "Impressao"
Option Explicit
Public ObjPreview       As Object
Public ConfigImp        As Object
Public CI               As CriteriosImpressao
Public ObjWord          As Object
Type CriteriosImpressao
    Qualidade   As String
    Orientacao  As String
    Fonte       As String
    tFonte      As String
    Negrito     As Boolean
    Italico     As Boolean
    Sublinhado  As Boolean
    Copias      As Integer
    Preview     As Boolean 'Informa se Preview (True) // Print (False)
End Type
Public Function cPreview(Optional Cab = 0, Optional strUE As String, Optional strUA As String)
    DoEvents
    Select Case Cab
        Case 0
            '
        Case 1
            Call CabImp1(strUE, strUA)
        Case 2
            Call CabImp2
        Case 3
            Call CabImp3(strUE, strUA)
        Case 4
            Call CabImp4(strUE, strUA)
        Case 5
            Call CabImp5(strUE, strUA)
        Case 6
            Call CabImp6(strUE, strUA)
    End Select
End Function
Public Function Relatorio(rpt As DataReport, sSQL As String, Optional Orientacao As Integer)

    AbrirBD_ADO
    Dim RsADO As ADODB.Recordset
    Set RsADO = New ADODB.Recordset
    
    RsADO.Open sSQL, BD_ADO
    Set rpt.DataSource = RsADO.DataSource
    
    Select Case Orientacao
        Case 1
            rpt.Orientation = rptOrientPortrait
        Case 2
            rpt.Orientation = rptOrientLandscape
        Case Else
            rpt.Orientation = rptOrientDefault
    End Select
    rpt.Sections("Rodape").Controls.Item("lbRodape").Caption = "CESNet  [Versão: " & VersaoAno & " (" & Versao & ")]"
    
End Function
Private Function CabImp1(cUE As String, cUA As String) ', cEnd As String, cTel As String)
'cUE - Unidade de Ensino
'cUA - Uniddae Adminsitrativa
'cEND - Enderço
'cTel - Telefone
 
'Dim fLarg As Integer
    'Set ObjPreview = Form_Preview.Pb_Folha
    ObjPreview.Font = "Arial"
    ObjPreview.PaintPicture MDIForm_Main.IL_LogoEstado.ListImages.Item(1).Picture, (ObjPreview.ScaleWidth / 2) - 510, 1300, 1020, 1020
    ObjPreview.FontSize = 12
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 2400 ' soma o ponto inicial da fig e o alt dela
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("GOVERNO DO ESTADO DO RIO DE JANEIRO") / 2)
    ObjPreview.Print "GOVERNO DO ESTADO DO RIO DE JANEIRO"
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("SECRETARIA DE ESTADO DE EDUCAÇÃO") / 2)
    ObjPreview.Print "SECRETARIA DE ESTADO DE EDUCAÇÃO"
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("SUBSECRETARIA DE EDUCAÇÃO") / 2)
    ObjPreview.Print "SUBSECRETARIA DE EDUCAÇÃO"
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("SUPERINTENDÊNCIA DE ENSINO") / 2)
    ObjPreview.Print "SUPERINTENDÊNCIA DE ENSINO"
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS") / 2)
    ObjPreview.Print "COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS"
    If Trim(cUE) = "" And Trim(cUA) = "" Then
        Else
            ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(cUE & " - " & cUA) / 2)
            ObjPreview.Print cUE & " - U.A.: " & cUA
    End If
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).AtoCriacao) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).AtoCriacao
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).AutorCurso) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).AutorCurso
    'If ObjPreview = "Printer" Then
    '    ObjPreview.EndDoc
    'End If
End Function

Private Function CabImp2()
    'Set ObjPreview = Form_Preview.Pb_Folha
    ObjPreview.Font = "Arial"
    ObjPreview.PaintPicture MDIForm_Main.IL_LogoEstado.ListImages.Item(1).Picture, 200, 200, 1020, 1020
    ObjPreview.FontSize = 8
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 200
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "GOVERNO DO ESTADO DO RIO DE JANEIRO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "SECRETARIA DE ESTADO DE EDUCAÇÃO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "SUBSECRETARIA DE EDUCAÇÃO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "SUPERINTENDÊNCIA DE ENSINO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS"
    ObjPreview.Print ""
    'If Trim(cUE) = "" And Trim(cUA) = "" Then
    '    Else
    '        ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(cUE & " - " & cUA) / 2)
    '        ObjPreview.Print cUE & " - " & cUA
    'End If
    'If ObjPreview = "Printer" Then
    '    ObjPreview.EndDoc
    'End If
End Function
Private Function CabImp3(cUE As String, cUA As String)
    'Set ObjPreview = Form_Preview.Pb_Folha
    ObjPreview.Font = "Arial"
    ObjPreview.PaintPicture MDIForm_Main.IL_LogoEstado.ListImages.Item(1).Picture, 200, 200, 1020, 1020
    ObjPreview.FontSize = 8
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 200
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "GOVERNO DO ESTADO DO RIO DE JANEIRO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "SECRETARIA DE ESTADO DE EDUCAÇÃO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "SUBSECRETARIA DE EDUCAÇÃO"
    ObjPreview.CurrentX = 1280
    ObjPreview.Print "SUPERINTENDÊNCIA DE ENSINO"
    ObjPreview.FontBold = True
    ObjPreview.CurrentX = 1280
    'ObjPreview.Print "COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS"
    ObjPreview.Print cUE & " - " & cUA
    'If Trim(cUE) = "" And Trim(cUA) = "" Then
    '    Else
    '        ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(cUE & " - " & cUA) / 2)
    '        ObjPreview.Print cUE & " - " & cUA
    'End If
    'If ObjPreview = "Printer" Then
    '    ObjPreview.EndDoc
    'End If
End Function
Private Function CabImp4(cUE As String, cUA As String)
    'Set ObjPreview = Form_Preview.Pb_Folha
    ObjPreview.Font = "Arial"
    'ObjPreview.PaintPicture MDIForm_Main.IL_LogoEstado.ListImages.Item(1).Picture, 200, 200, 1020, 1020
    ObjPreview.FontSize = 8
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 200
    ObjPreview.CurrentX = 250
    ObjPreview.Print "GOVERNO DO ESTADO DO RIO DE JANEIRO"
    ObjPreview.CurrentX = 250
    ObjPreview.Print "SECRETARIA DE ESTADO DE EDUCAÇÃO"
    ObjPreview.CurrentX = 250
    ObjPreview.Print "SUBSECRETARIA DE EDUCAÇÃO"
    ObjPreview.CurrentX = 250
    ObjPreview.Print "SUPERINTENDÊNCIA DE ENSINO"
    ObjPreview.FontBold = True
    ObjPreview.CurrentX = 250
    'ObjPreview.Print "COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS"
    ObjPreview.Print cUE & " - " & cUA
    'If Trim(cUE) = "" And Trim(cUA) = "" Then
    '    Else
    '        ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(cUE & " - " & cUA) / 2)
    '        ObjPreview.Print cUE & " - " & cUA
    'End If
    'If ObjPreview = "Printer" Then
    '    ObjPreview.EndDoc
    'End If
End Function
Private Function CabImp5(cUE As String, cUA As String)
    'Set ObjPreview = Form_Preview.Pb_Folha
    ObjPreview.Font = "Arial"
    'ObjPreview.PaintPicture MDIForm_Main.IL_LogoEstado.ListImages.Item(1).Picture, 200, 200, 1020, 1020
    ObjPreview.FontSize = 10
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 800
    ObjPreview.CurrentX = 650
    ObjPreview.Print "[ FOTO ]"
    ObjPreview.Line (170, 170)-(1800, 2000), , B
    
    ObjPreview.FontSize = 20
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 200
    ObjPreview.CurrentX = 2000
    ObjPreview.Print "CEJA - Centro de Estudos de Jovens e Adultos"
    

    ObjPreview.FontSize = 8
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 600
    ObjPreview.CurrentX = 2000
    ObjPreview.Print "GOVERNO DO ESTADO DO RIO DE JANEIRO"
    ObjPreview.CurrentX = 2000
    ObjPreview.Print "SECRETARIA DE ESTADO DE EDUCAÇÃO"
    ObjPreview.CurrentX = 2000
    ObjPreview.Print "SUBSECRETARIA DE EDUCAÇÃO"
    ObjPreview.CurrentX = 2000
    ObjPreview.Print "SUPERINTENDÊNCIA DE ENSINO"
    ObjPreview.FontBold = True
    ObjPreview.CurrentX = 2000
    'ObjPreview.Print "COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS"
    ObjPreview.Print cUE & " - " & cUA
    'If Trim(cUE) = "" And Trim(cUA) = "" Then
    '    Else
    '        ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(cUE & " - " & cUA) / 2)
    '        ObjPreview.Print cUE & " - " & cUA
    'End If
    'If ObjPreview = "Printer" Then
    '    ObjPreview.EndDoc
    'End If
End Function
Private Function CabImp6(cUE As String, cUA As String)
    
    ObjPreview.Font = "Arial"
    ObjPreview.Line (100, 100)-(11700, 4000), , B 'Linha Externa
    ObjPreview.Line (250, 250)-(11500, 3900), , B 'Linha Interna
    ObjPreview.Line (5800, 250)-(5800, 3950) 'Linha de divisao
    ObjPreview.Line (350, 1800)-(1900, 3780), , B 'Retangulo Foto
    ObjPreview.PaintPicture MDIForm_Main.IL_LogoEstado.ListImages.Item(1).Picture, 400, 300, 1020, 1020
    ObjPreview.FontSize = 12
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 400
    ObjPreview.CurrentX = 1500
    ObjPreview.Print "CEJA - Centro de Estudos de Jovens e Adultos"
    

    ObjPreview.FontSize = 8
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 680
    ObjPreview.CurrentX = 1500
    ObjPreview.Print "GOVERNO DO ESTADO DO RIO DE JANEIRO"
    ObjPreview.CurrentX = 1500
    ObjPreview.Print "SECRETARIA DE ESTADO DE EDUCAÇÃO"
    'ObjPreview.CurrentX = 1500
    'ObjPreview.Print "SUBSECRETARIA DE EDUCAÇÃO"
    'ObjPreview.CurrentX = 1500
    'ObjPreview.Print "SUPERINTENDÊNCIA DE ENSINO"
    ObjPreview.FontBold = True
    ObjPreview.CurrentX = 1500
    'ObjPreview.Print "COORDENAÇÃO DE ENSINO DE JOVENS E ADULTOS"
    ObjPreview.FontSize = 10
    ObjPreview.Print cUE & " - U.A.:" & cUA
    ObjPreview.FontSize = 8
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 2700
    ObjPreview.CurrentX = 700
    ObjPreview.Print "[ F O T O ]"
End Function
Public Sub CabImp7(Titulo As String)
    Dim v As Integer
    Dim h As Integer
    Dim NomeSistema As String
    NomeSistema = "CESNet"
    v = 1200
    h = 500
    
    Printer.ScaleMode = 1
    
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
    
    Printer.FontSize = 14
    Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(NomeSistema) / 2)
    Printer.CurrentY = 600
    Printer.Print NomeSistema
    
    Printer.FontSize = 10
    
    Printer.Line (h, v)-(Printer.ScaleWidth - 500, v)
    
    Printer.CurrentX = (Printer.ScaleWidth / 2) - (Printer.TextWidth(Titulo) / 2)
    Printer.CurrentY = 1300
    Printer.Print Titulo
    
    v = v + Printer.TextHeight(Titulo) + 150
    Printer.Line (h, v)-(Printer.ScaleWidth - 500, v)
    
    Printer.Print
    Printer.FontSize = 10
    Printer.FontBold = False
    Printer.FontItalic = False
    Printer.FontUnderline = False
End Sub

Public Sub ImpRodape()
    ObjPreview.CurrentY = ObjPreview.ScaleHeight - 400
    ObjPreview.Print Tab(10); "CESNet [v." & Versao & "]"
End Sub

Public Function ImpFichaMatr(MatrID As String)
   'On Error GoTo msg
   Dim cIni As Integer
    Call cPreview(5, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Print
    ObjPreview.FontSize = 16
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("FICHA DE MATRICULA") / 2)
    ObjPreview.Print "FICHA DE MATRICULA"
    
    ObjPreview.Print
    ObjPreview.FontSize = 10
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 3000 'vertical
    ObjPreview.CurrentX = Printer.ScaleWidth - 3000
    ObjPreview.Print "Matricula:"
    
    ObjPreview.FontSize = 14
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.CurrentY = 2945 'vertical
    ObjPreview.CurrentX = Printer.ScaleWidth - 2000 '1550
    ObjPreview.Print PgDadosMatr(MatrID).MatrID
    
    ObjPreview.FontSize = 10
    ObjPreview.FontBold = False
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    'ObjPreview.CurrentY = 3000 'vertical
    ObjPreview.CurrentX = Printer.ScaleWidth - 2500
    ObjPreview.Print "Data:"

    ObjPreview.FontSize = CI.tFonte
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontItalic = CI.Italico
    ObjPreview.FontUnderline = CI.Sublinhado
    ObjPreview.CurrentY = 3290 'vertical
    ObjPreview.CurrentX = Printer.ScaleWidth - 1550
    ObjPreview.Print PgDadosMatr(MatrID).DtMatr
    
    ObjPreview.Print Tab(10); "Unidade:"; Tab(30); PgDadosMatr(MatrID).UnidMatr
    
    ObjPreview.Print Tab(10); "Aluno(a):"; Tab(30); PgDadosMatr(MatrID).Nome
    ObjPreview.Print Tab(10); "Endereço:"; Tab(30); PgDadosMatr(MatrID).Endereco & "   Num.: " & PgDadosMatr(MatrID).Numero & "   Compl.: " & PgDadosMatr(MatrID).Compl
    ObjPreview.Print Tab(30); PgDadosMatr(MatrID).Bairro
    ObjPreview.Print Tab(30); PgDadosMatr(MatrID).Munic & " / " & PgDadosMatr(MatrID).UF
    ObjPreview.Print Tab(30); PgDadosMatr(MatrID).CEP
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print Tab(10); "E-Mail:"; Tab(30); PgDadosMatr(MatrID).Mail
    ObjPreview.Print Tab(10); "Cel.:"; Tab(30); PgDadosMatr(MatrID).Cel
    ObjPreview.Print Tab(10); "Tel.(s):"; Tab(30); PgDadosMatr(MatrID).Tel1
    ObjPreview.Print Tab(30); PgDadosMatr(MatrID).Tel2
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print Tab(10); "R.G.:"; Tab(30); PgDadosMatr(MatrID).RG
    ObjPreview.Print Tab(10); "Orgão Emissor:"; Tab(30); PgDadosMatr(MatrID).OE
    ObjPreview.Print Tab(10); "CPF:"; Tab(30); PgDadosMatr(MatrID).CPF
    ObjPreview.Print Tab(10); "Sexo:"; Tab(30); PgDadosMatr(MatrID).Sexo
    ObjPreview.Print Tab(10); "Est. Civil:"; Tab(30); PgDadosMatr(MatrID).EstCivil
    ObjPreview.Print Tab(10); "Nacionalidade:"; Tab(30); PgDadosMatr(MatrID).Nacion
    ObjPreview.Print Tab(10); "Natural:"; Tab(30); PgDadosMatr(MatrID).Natural
    ObjPreview.Print Tab(10); "Nascimento:"; Tab(30); PgDadosMatr(MatrID).Nasc
    
    ObjPreview.Print Tab(10); "Deficiente:"; Tab(30); IIf(PgDadosMatr(MatrID).Deficiencia = 0, "", PgDadosMatr(MatrID).Deficiencia)
    ObjPreview.Print
    ObjPreview.Print Tab(10); "Pai:"; Tab(30); PgDadosMatr(MatrID).Pai
    ObjPreview.Print Tab(10); "Mãe"; Tab(30); PgDadosMatr(MatrID).Mae
    ObjPreview.Print
    cIni = 1
    ObjPreview.Print Tab(10); "Obs:";
    Do Until Trim(Mid(UCase(PgDadosMatr(MatrID).Obs), cIni, 100)) = ""
        ObjPreview.Print Tab(15); Trim(Mid(UCase(PgDadosMatr(MatrID).Obs), cIni, 100))
        cIni = cIni + 100
    Loop
    
    ObjPreview.CurrentY = ObjPreview.ScaleHeight - 2500
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).Municipio & ", " & Date) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & Date
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosMatr(MatrID).Nome) / 2)
    ObjPreview.Print PgDadosMatr(MatrID).Nome
    
    Call ImpRodape
    
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
    '*************************************************
   
   
    Exit Function
msg:
    Call RegLogErros(Err.Number, Err.Description, "Impressao Ficha de Matricula", UsuarioID)
    MsgBox Err.Description, vbInformation, Err.Number
    Exit Function
End Function
Public Sub ImpCarteirinha(MatrID As String)
    ''on error goto msg
    'Call cPreview(6, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
   '
   ' ObjPreview.Font = CI.Fonte
   ' ObjPreview.CurrentY = 2000
   '
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print Tab(29); "Matricula:"
   ' ObjPreview.FontSize = 14
   ' ObjPreview.FontBold = True
   ' 'ObjPreview.FontName = "barcode font"
   ' ObjPreview.Print Tab(18); PgDadosMatr(MatrID).MatrID
   ' ObjPreview.FontName = CI.Fonte
   '
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print Tab(29); "Aluno(a):"
   ' ObjPreview.FontSize = 10
   ' ObjPreview.FontBold = True
   ' If Len(PgDadosMatr(MatrID).Nome) > 26 Then
   '         ObjPreview.Print Tab(26); Trim(Mid(PgDadosMatr(MatrID).Nome, 1, 26))
   '         ObjPreview.Print Tab(26); Trim(Mid(PgDadosMatr(MatrID).Nome, 27))
   '     Else
   '         ObjPreview.Print Tab(26); PgDadosMatr(MatrID).Nome
   ' End If
   '
   ' ObjPreview.CurrentY = 3300
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print Tab(32); "Data de Nasc.:"
   ' ObjPreview.FontSize = 10
   ' 'ObjPreview.FontBold = True
   ' ObjPreview.Print Tab(28); PgDadosMatr(MatrID).Nasc
   '
   ' ObjPreview.CurrentY = 3300
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print Tab(52); "Validade da Carteirinha:"
   ' ObjPreview.FontSize = 12
   ' ObjPreview.FontBold = True
   ' ObjPreview.Print Tab(35); PgDadosMatr(MatrID).ValCard
    
    
   ' ObjPreview.CurrentY = 450
    
    'ObjPreview.FontSize = 8
    'ObjPreview.FontBold = False
    'ObjPreview.Print Tab(84); "Unidade Administrativa:"
    'ObjPreview.FontSize = 10
    'ObjPreview.FontBold = True
    'ObjPreview.Print Tab(72); PgDadosMatr(MatrID).UnidMatr
    
    
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' 'ObjPreview.Print
   ' ObjPreview.Print Tab(84); "Filiação:"
   ' ObjPreview.FontBold = True
   ' ObjPreview.Print Tab(84); PgDadosMatr(MatrID).Pai
   ' ObjPreview.Print Tab(84); PgDadosMatr(MatrID).Mae
   ' ObjPreview.Print
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print Tab(84); "Curso:"
   ' ObjPreview.FontBold = True
   ' ObjPreview.Print Tab(84); PgNomeEnsino(PgMatrEnsino(MatrID, False))
   ' ObjPreview.Print
    
    
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print Tab(84); "Obs:"
   ' ObjPreview.FontBold = True
   ' 'ObjPreview.Print Tab(84); "TODA E QUALQUER ATIVIDADE NO CES SÓ PODERÁ SER EFETUADA"
   ' 'ObjPreview.Print Tab(84); "MEDIANTE A APRESENTAÇÃO DESTA CARTEIRA."
   ' 'ObjPreview.Print Tab(84); "O ALUNO DEVERÁ ESTAR UNIFORMIZADO E APRESENTAR ESTA"
   ' ObjPreview.Print Tab(84); "O ALUNO DEVERÁ APRESENTAR ESTA CARTEIRA ESCOLAR"
   ' ObjPreview.Print Tab(84); "PARA QUALQUER ATIVIDADE NO CES."
   '
   ' ObjPreview.FontSize = 8
   ' ObjPreview.FontBold = False
   ' ObjPreview.Print
   ' ObjPreview.Print
    'ObjPreview.Print
   ' ObjPreview.Print Tab(85); "___________________________    ___________________________"
   ' ObjPreview.Print Tab(85); "                  Aluno(a)                                            Diretor(a)"
    
    'ObjPreview.Print
    'ObjPreview.Print
    'ObjPreview.Print Tab(85); "_____________________________"
    'ObjPreview.Print Tab(97); "Diretor(a)"
    
    'Call CodBar39(MatrID, ObjPreview, 500, 200, 50) =>> Erro na hora da impressao
    'Codigo de barras
    'ObjPreview.CurrentY = 3250 '2900
    'ObjPreview.FontSize = 30  '45
    'ObjPreview.FontBold = False
    'ObjPreview.FontName = "c39hrp36dltt"  '"barcode font"
    '67
    'ObjPreview.Print Tab(32); Left(PgDadosMatr(MatrID).MatrID, 2) & Mid(PgDadosMatr(MatrID).MatrID, 4, 3) & Right(PgDadosMatr(MatrID).MatrID, 4)
    'ObjPreview.FontName = CI.Fonte
    
    'If CI.Preview = False Then
    '    ObjPreview.EndDoc
    'End If
   '
   ' Exit Sub
'msg:
'    Call RegLogErros(Err.Number, Err.Description, "Impressao Ficha de Matricula", UsuarioID)
'    MsgBox Err.Description, vbInformation, Err.Number
'    Exit Sub
End Sub
Public Sub RelatorioWord(Doc As String, MatrID As String, nCurso As String, DataHoje As String)
    On Error GoTo trata_erro
    Dim idCurso As Integer
    Dim caminho As String
    
    caminho = PathBD & "\Database\RptModelo\" & Doc
    
    If Dir(caminho) = "" Then
        MsgBox "Erro ao localizar arquivo MODELO do documento." & vbCrLf & "Impressão cancelada!", vbCritical, "Aviso"
        Exit Sub
    End If
    
    idCurso = PgIDEnsino(nCurso)
    Set ObjWord = New Word.Application
    ObjWord.Visible = False
    ObjWord.Documents.Open (caminho)

    ' chama rotina para substituicao
    Call Substitui_Var("@matricula", MatrID)
    Call Substitui_Var("@datamatr", PgDadosMatr(MatrID).DtMatr)
    Call Substitui_Var("@dataret", PgDadosMatr(MatrID).DtRetorno)
    
    Call Substitui_Var("@unidmatr", PgDadosMatr(MatrID).UnidMatr)
    
    Call Substitui_Var("@nome", PgDadosMatr(MatrID).Nome)
    Call Substitui_Var("@endereco", PgDadosMatr(MatrID).Endereco)
    Call Substitui_Var("@numero", PgDadosMatr(MatrID).Numero)
    Call Substitui_Var("@complemento", PgDadosMatr(MatrID).Compl)
    Call Substitui_Var("@bairro", PgDadosMatr(MatrID).Bairro)
    Call Substitui_Var("@municipio", PgDadosMatr(MatrID).Munic)
    Call Substitui_Var("@uf", PgDadosMatr(MatrID).UF)
    Call Substitui_Var("@cep", PgDadosMatr(MatrID).CEP)
    
    Call Substitui_Var("@sexo", PgDadosMatr(MatrID).Sexo)
    
    Call Substitui_Var("@mail", PgDadosMatr(MatrID).Mail)
    Call Substitui_Var("@celular", PgDadosMatr(MatrID).Cel)
    Call Substitui_Var("@telefone1", PgDadosMatr(MatrID).Tel1)
    Call Substitui_Var("@telefone2", PgDadosMatr(MatrID).Tel2)
    
    Call Substitui_Var("@cpf", PgDadosMatr(MatrID).CPF)
    Call Substitui_Var("@rg", PgDadosMatr(MatrID).RG)
    Call Substitui_Var("@orgemissor", PgDadosMatr(MatrID).OE)
    Call Substitui_Var("@certnascimento", PgDadosMatr(MatrID).CertNasc)
    
    Call Substitui_Var("@mae", PgDadosMatr(MatrID).Mae)
    Call Substitui_Var("@pai", PgDadosMatr(MatrID).Pai)
    
    Call Substitui_Var("@nascimento", PgDadosMatr(MatrID).Nasc)
    Call Substitui_Var("@nacionalidade", PgDadosMatr(MatrID).Nacion)
    Call Substitui_Var("@naturaluf", PgDadosMatr(MatrID).NaturalUF)
    Call Substitui_Var("@natural", PgDadosMatr(MatrID).Natural)
    Call Substitui_Var("@estadocivil", PgDadosMatr(MatrID).EstCivil)
    Call Substitui_Var("@deficiencia", PgNomeDef(PgDadosMatr(MatrID).Deficiencia))
    
    Call Substitui_Var("@validcarteira", CStr(PgDadosMatr(MatrID).ValCard))
    
    Call Substitui_Var("@numanterior", PgDadosMatr(MatrID).NumAnt)
    Call Substitui_Var("@obs", PgDadosMatr(MatrID).Obs)
    
    Call Substitui_Var("@statusmatr", PgStatusMatricula(MatrID))
    
    Call Substitui_Var("@datahoje", DataHoje)
    
    Call Substitui_Var("@curso_inicio", PgDadosCurso(MatrID, idCurso).DtInicio)
    Call Substitui_Var("@curso_final", PgDadosCurso(MatrID, idCurso).DtFinal)
    Call Substitui_Var("@curso_local", PgDadosCurso(MatrID, idCurso).Local)
    Call Substitui_Var("@curso", nCurso)
    Dim Rst      As Recordset
    Dim sSQL     As String
    
    
    sSQL = "SELECT MatriculaDisciplina.MatrID, MatriculaDisciplina.EnsinoID, Disciplina.Sigla, Disciplina.Descr, MatriculaDisciplina.DtInicio, MatriculaDisciplina.DtConclusao, MatriculaDisciplina.Local, MatriculaDisciplina.Cidade, MatriculaDisciplina.UF " & _
           "FROM MatriculaDisciplina INNER JOIN Disciplina ON MatriculaDisciplina.DisciplinaID = Disciplina.ID " & _
           "WHERE (((MatriculaDisciplina.MatrID)= '" & MatrID & "') AND ((MatriculaDisciplina.EnsinoID)=" & idCurso & ") AND ((Disciplina.Sigla) Is Not Null))"

    
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
        Else
            Rst.MoveFirst
            Do Until Rst.EOF
                Call Substitui_Var("@disc_" & Rst.Fields("Sigla") & "_inicio", IIf(IsNull(Rst.Fields("DtInicio")), "", Rst.Fields("DtInicio")))
                Call Substitui_Var("@disc_" & Rst.Fields("Sigla") & "_final", IIf(IsNull(Rst.Fields("DtConclusao")), "", Rst.Fields("DtConclusao")))
                Call Substitui_Var("@disc_" & Rst.Fields("Sigla") & "_local", IIf(IsNull(Rst.Fields("Local")), "", Rst.Fields("Local")))
                Call Substitui_Var("@disc_" & Rst.Fields("Sigla") & "_cidade", IIf(IsNull(Rst.Fields("Cidade")), "", Rst.Fields("Cidade")))
                Call Substitui_Var("@disc_" & Rst.Fields("Sigla") & "_uf", IIf(IsNull(Rst.Fields("UF")), "", Rst.Fields("UF")))
                Rst.MoveNext
            Loop
    End If
    
    
    'Salva o documento com um novo nome
    ObjWord.ActiveDocument.SaveAs (PathBD & "\Database\RptTMP\" & Format(MatrID, "000000000") & "-" & Format(Date, "DDMMYYYY") & Format(Time, "HHMMSS") & ".doc")
    'ObjWord.ActiveDocument.noSave
    ObjWord.Visible = False
    ObjWord.PrintOut
    'Encerra o word
    ObjWord.Quit
    ' informa ao usuario que o contrato foi gerado
    'MsgBox "Contrato gerado com sucesso! em : " '& txtcontrato, vbInformation, " Contrato Gerado "
    ' libera memoria
    Set ObjWord = Nothing
    Exit Sub
trata_erro:
    MsgBox "Ocorreu um erro durante o processamento." & vbCrLf & vbCrLf & "- Erro numero: " & Err.Number & vbCrLf & vbCrLf & "- Descrição: " & Err.Description
    'ObjWord.Quit
    Resume Next


End Sub

Private Sub Substitui_Var(Header As String, Data As String)
    On Error Resume Next
'    With ObjWord.Selection.Find
'        .ClearFormatting
'        .Text = Header
'        .Execute Forward:=True
'    End With
'    Clipboard.Clear
'    Clipboard.SetText (Data)
'    ObjWord.Selection.Paste
'    Clipboard.Clear

    ObjWord.Selection.Find.ClearFormatting
    ObjWord.Selection.Find.Replacement.ClearFormatting
    With ObjWord.Selection.Find
        .Text = Header
        .Replacement.Text = Data
        .Forward = True
        .Wrap = wdFindContinue
        .Format = False
        .MatchCase = False
        .MatchWholeWord = False
        .MatchWildcards = False
        .MatchSoundsLike = False
        .MatchAllWordForms = False
    End With
    ObjWord.Selection.Find.Execute Replace:=wdReplaceAll


End Sub

Public Sub Rpt001(MatrID As String, EnsinoID As Integer, DtDoc As String)

'*****************************************************
   'Call RelatorioWord("001", "08.001.0001", "Fundamental", "21/10/2009")
    'Exit Sub
    
    
    
    
    Dim strNome As String
    Dim Ensino As String
    Ensino = PgNomeEnsino(EnsinoID)
    Ensino = IIf(PgNomeEnsino(PgMatrEnsino(MatrID)) = 0, "NENHUM ENSINO", PgNomeEnsino(PgMatrEnsino(MatrID)))
    If Ensino = "NENHUM ENSINO" Then
        MsgBox "Esta matricula (" & MatrID & ") não possue ensino iniciado.", vbInformation, "CESNet - Aviso"
        Exit Sub
    End If
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
    Call cPreview(1, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.FontSize = 18
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("D E C L A R A Ç Ã O") / 2)
    ObjPreview.CurrentY = 6000 'vertical
    
    
    ObjPreview.Print "D E C L A R A Ç Ã O"
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontSize = CI.tFonte
    
    
    'Ensino = Ensino & "  " & Mid(String(19, "X"), 1, 20 - Len(Ensino))
    'strNome = Nome '& "  " & Mid(String(49, "X"), 1, 50 - Len(Nome)) & ","
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
      
    ObjPreview.Print Tab(10); "Matrícula: "; Tab(25); MatrID
    ObjPreview.Print Tab(10); "Aluno(a): "; Tab(25); PgDadosMatr(MatrID).Nome
    'ObjPreview.Print Tab(10); "Endereço: "; Tab(25); PgDadosMatr(MatrID).Endereco
    'ObjPreview.Print Tab(10); "Bairro: "; Tab(25); PgDadosMatr(MatrID).Bairro
    'ObjPreview.Print Tab(10); "Município: "; Tab(25); PgDadosMatr(MatrID).Munic; " / " & PgDadosMatr(MatrID).UF
    'ObjPreview.Print Tab(10); "CEP: "; Tab(25); PgDadosMatr(MatrID).CEP
    'ObjPreview.Print Tab(10); "Nascimento: "; Tab(25); PgDadosMatr(MatrID).Nasc
    ObjPreview.Print Tab(10); "RG: "; Tab(25); PgDadosMatr(MatrID).RG & " (" & PgDadosMatr(MatrID).OE & ")"
    'ObjPreview.Print Tab(10); "Naturalidade: "; Tab(25); PgDadosMatr(MatrID).Natural
    'ObjPreview.Print Tab(10); "Filiação: "; Tab(25); PgDadosMatr(MatrID).Pai
    'ObjPreview.Print Tab(25); PgDadosMatr(MatrID).Mae
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.Print Tab(15); "Declaro que o(a) aluno(a) acima está cursando o Ensino " & Ensino & ", neste estabelecimento de ensino,"
    ObjPreview.Print Tab(8); "que utiliza metodologia de ensino modular e individualizado sem caráter de série."
    'ObjPreview.Print Tab(8); ""
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).UA & ", " & DtDoc) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    'ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(UnidadeEnsinoNome) / 2)
    'ObjPreview.Print UnidadeEnsinoNome
    Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If

End Sub

Public Sub Rpt002(MatrID As String, EnsinoID As Integer, DtDoc As String)
    Dim RsMatriculaDisciplina As Recordset
    'Dim EnsinoID As Integer
    'EnsinoID = PgMatrEnsino(MebMatricula.Text, True)
    'If EnsinoID = 0 Then
    '    MsgBox "Esta Matricula não possue Ensino concluido. Tente tirar uma declaração parcial.", vbInformation, "CESNet - Aviso!"
    '    Exit Sub
    'End If
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            MsgBox "Esta Matricula não possue DISCIPLINA(S) cadastrada(s)!", vbInformation, "CESNet - Atenção"
        Else
            RsMatriculaDisciplina.MoveFirst
    End If
    Call cPreview(1, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.FontSize = 18
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("D E C L A R A Ç Ã O  D E  C O N C L U S Ã O") / 2)
    ObjPreview.CurrentY = 5000 'vertical
    
    
    ObjPreview.Print "D E C L A R A Ç Ã O  D E  C O N C L U S Ã O"
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontSize = CI.tFonte
    
    Dim strNome As String
    Dim Ensino As String
    
    Ensino = IIf(PgNomeEnsino(PgMatrEnsino(MatrID)) = 0, "NENHUM ENSINO", PgNomeEnsino(PgMatrEnsino(MatrID)))
    'Ensino = Ensino & "  " & Mid(String(19, "X"), 1, 20 - Len(Ensino))
    'strNome = Nome '& "  " & Mid(String(49, "X"), 1, 50 - Len(Nome)) & ","
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
      
    ObjPreview.Print Tab(10); "Matrícula: "; Tab(25); PgDadosMatr(MatrID).MatrID
    ObjPreview.Print Tab(10); "Aluno(a): "; Tab(25); PgDadosMatr(MatrID).Nome
    ObjPreview.Print Tab(10); "Endereço: "; Tab(25); PgDadosMatr(MatrID).Endereco
    ObjPreview.Print Tab(10); "Bairro: "; Tab(25); PgDadosMatr(MatrID).Bairro
    ObjPreview.Print Tab(10); "Município: "; Tab(25); PgDadosMatr(MatrID).Munic; " / " & PgDadosMatr(MatrID).UF
    ObjPreview.Print Tab(10); "CEP: "; Tab(25); PgDadosMatr(MatrID).CEP
    ObjPreview.Print Tab(10); "Nascimento: "; Tab(25); PgDadosMatr(MatrID).Nasc
    ObjPreview.Print Tab(10); "RG: "; Tab(25); PgDadosMatr(MatrID).RG & " (" & PgDadosMatr(MatrID).OE & ")"
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    'ObjPreview.Print
    
    ObjPreview.Print Tab(15); "Declaramos para os devidos fins de escolaridade, que o(a) aluno(a) acima concluiu o ensino " & PgNomeEnsino(EnsinoID)
    ObjPreview.Print Tab(8); "nesta institução de ensino, apresentando o seguinte quadro de avaliação:"
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(10); "DISCIPLINA"; Tab(40); "DATA DA CONCLUSÃO"; Tab(68); "LOCAL"
    ObjPreview.FontBold = CI.Negrito
    Do Until RsMatriculaDisciplina.EOF
        ObjPreview.Print Tab(10); Trim(PgNomeDisciplina(RsMatriculaDisciplina.Fields("DisciplinaID"))); _
                         Tab(45); IIf(IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) = True, "00/00/0000", RsMatriculaDisciplina.Fields("DtConclusao")); _
                         Tab(68); IIf(IsNull(RsMatriculaDisciplina.Fields("Local")), "", RsMatriculaDisciplina.Fields("Local"))
        RsMatriculaDisciplina.MoveNext
    Loop
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    'ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(UnidadeEnsinoNome) / 2)
    'ObjPreview.Print UnidadeEnsinoNome
    Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If

End Sub

Public Sub Rpt003(MatrID As String, EnsinoID As Integer, DtDoc As String)
    Dim RsMatriculaDisciplina As Recordset
    'Dim EnsinoID As Integer
    'EnsinoID = PgMatrEnsino(MatrID, False)
    'If EnsinoID = 0 Then
    '    MsgBox "Esta Matricula não possue Ensino iniciado. Tente tirar uma declaração total.", vbInformation, "CESNet - Aviso!"
    '    Exit Sub
    'End If
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DtConclusao <> Null")
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            MsgBox "Esta Matricula não possue DISCIPLINA(S) cadastrada(s)!", vbInformation, "CESNet - Atenção"
            Exit Sub
        Else
            RsMatriculaDisciplina.MoveFirst
    End If
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
    
    
    Call cPreview(1, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.FontSize = 18
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("D E C L A R A Ç Ã O  D E  C O N C L U S Ã O  P A R C I A L") / 2)
    ObjPreview.CurrentY = 6000 'vertical
    
    
    ObjPreview.Print "D E C L A R A Ç Ã O  D E  C O N C L U S Ã O  P A R C I A L"
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontSize = CI.tFonte
    
    Dim strNome As String
    Dim Ensino As String
    
    Ensino = IIf(PgNomeEnsino(PgMatrEnsino(MatrID)) = 0, "NENHUM ENSINO", PgNomeEnsino(PgMatrEnsino(MatrID)))
    'Ensino = Ensino & "  " & Mid(String(19, "X"), 1, 20 - Len(Ensino))
    'strNome = Nome '& "  " & Mid(String(49, "X"), 1, 50 - Len(Nome)) & ","
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
      
    ObjPreview.Print Tab(10); "Matrícula: "; Tab(25); PgDadosMatr(MatrID).MatrID
    ObjPreview.Print Tab(10); "Aluno(a): "; Tab(25); PgDadosMatr(MatrID).Nome
    ObjPreview.Print Tab(10); "Endereço: "; Tab(25); PgDadosMatr(MatrID).Endereco
    ObjPreview.Print Tab(10); "Bairro: "; Tab(25); PgDadosMatr(MatrID).Bairro
    ObjPreview.Print Tab(10); "Município: "; Tab(25); PgDadosMatr(MatrID).Munic; " / " & PgDadosMatr(MatrID).UF
    ObjPreview.Print Tab(10); "CEP: "; Tab(25); PgDadosMatr(MatrID).CEP
    ObjPreview.Print Tab(10); "Nascimento: "; Tab(25); PgDadosMatr(MatrID).Nasc
    ObjPreview.Print Tab(10); "RG: "; Tab(25); PgDadosMatr(MatrID).RG & " (" & PgDadosMatr(MatrID).OE & ")"
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.Print Tab(15); "Declaramos para os devidos fins de escolaridade, que o(a) aluno(a) acima está devidamente"
    ObjPreview.Print Tab(8); "matriculado nesta instituição de ensino e apresenta o seguinte quadro de avaliação no ensino " & PgNomeEnsino(EnsinoID) & ":"
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(10); "DISCIPLINA"; Tab(40); "DATA DA CONCLUSÃO"; Tab(90); "LOCAL"
    ObjPreview.FontBold = CI.Negrito
    Do Until RsMatriculaDisciplina.EOF
        ObjPreview.Print Tab(10); Trim(PgNomeDisciplina(RsMatriculaDisciplina.Fields("DisciplinaID"))); _
                         Tab(46); IIf(IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) = True, "00/00/0000", RsMatriculaDisciplina.Fields("DtConclusao")); _
                         Tab(75); IIf(IsNull(RsMatriculaDisciplina.Fields("Local")), "", RsMatriculaDisciplina.Fields("Local"))
        RsMatriculaDisciplina.MoveNext
    Loop
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    'ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(UnidadeEnsinoNome) / 2)
    'ObjPreview.Print UnidadeEnsinoNome
    Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If

End Sub
Public Sub Rpt004(MatrID As String, DtDoc As String)
    Dim RsMatriculaProva As Recordset
    Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND DtAvaliacao = #" & Format(DtDoc, "MM/DD,YYYY") & "#")
    If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
        If MsgBox("Esta matricula não possue avaliação na presente data (" & DtDoc & ")." & Chr(13) & "Gostaria de emitir assim mesmo a DECLARAÇÃO?", vbInformation + vbYesNo, "CESNet - Aviso") = vbNo Then
            Exit Sub
        End If
    End If
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
    Call cPreview(1, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.FontSize = 18
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("D E C L A R A Ç Ã O") / 2)
    ObjPreview.CurrentY = 6000 'vertical
    
    
    ObjPreview.Print "D E C L A R A Ç Ã O"
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontSize = CI.tFonte
    
    Dim strNome As String
    Dim Ensino As String
    
    'Ensino = PgNomeEnsino(EnsinoID)
    'Ensino = IIf(PgNomeEnsino(PgMatrEnsino(MatrID)) = 0, "NENHUM ENSINO", PgNomeEnsino(PgMatrEnsino(MatrID)))
    'Ensino = Ensino & "  " & Mid(String(19, "X"), 1, 20 - Len(Ensino))
    'strNome = Nome '& "  " & Mid(String(49, "X"), 1, 50 - Len(Nome)) & ","
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
      
    
    ObjPreview.Print Tab(10); "Aluno(a): "; Tab(25); PgDadosMatr(MatrID).Nome
    ObjPreview.Print Tab(10); "Endereço: "; Tab(25); PgDadosMatr(MatrID).Endereco
    ObjPreview.Print Tab(10); "Bairro: "; Tab(25); PgDadosMatr(MatrID).Bairro
    ObjPreview.Print Tab(10); "Município: "; Tab(25); PgDadosMatr(MatrID).Munic; " / " & PgDadosMatr(MatrID).UF
    ObjPreview.Print Tab(10); "CEP: "; Tab(25); PgDadosMatr(MatrID).CEP
    ObjPreview.Print Tab(10); "Nascimento: "; Tab(25); PgDadosMatr(MatrID).Nasc
    ObjPreview.Print Tab(10); "RG: "; Tab(25); PgDadosMatr(MatrID).RG & " (" & PgDadosMatr(MatrID).OE & ")"
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    
    ObjPreview.Print Tab(15); "Declaro para os devidos fins que o(a) aluno(a) acima estando devidamente matriculado(a)"
    ObjPreview.Print Tab(8); "nesta Unidade de Ensino, compareceu nesta data para efetuar prova(s)."
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(UnidadeEnsinoNome) / 2)
    ObjPreview.Print UnidadeEnsinoNome
    Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If

End Sub

Public Sub Rpt005(MatrID As String, EnsinoID As Integer, DtDoc As String)

Dim Criterio As String

    Criterio = "SELECT MatriculaProva.*, MatriculaProva.MatrID, Disciplina.Descr, MatriculaProva.EnsinoID " & _
                "FROM MatriculaProva INNER JOIN Disciplina ON MatriculaProva.DisciplinaID =  Disciplina.ID " & _
                "WHERE (((MatriculaProva.MatrID)='" & MatrID & "') AND ((MatriculaProva.EnsinoID)=" & EnsinoID & ")) ORDER BY MatriculaProva.DisciplinaID, MatriculaProva.NProva"


    Call Relatorio(rptMatrLstProvas, Criterio)
    
    rptMatrLstProvas.Sections("Cab").Controls.Item("lbData").Caption = DtDoc
    rptMatrLstProvas.Sections("Cab").Controls.Item("lbMatricula").Caption = MatrID
    rptMatrLstProvas.Sections("Cab").Controls.Item("lbNome").Caption = PgDadosMatr(MatrID).Nome
    rptMatrLstProvas.Sections("Cab").Controls.Item("lbCurso").Caption = PgNomeEnsino(EnsinoID)
    rptMatrLstProvas.Show 1




Exit Sub
    
    
    
    
    Dim RsMatrProva As Recordset
    
    Dim DisciplinaID As Integer
    Dim EnsinoIndice As Integer 'Sera as provas referente ao ensino listado
   

    EnsinoIndice = EnsinoID
    Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' ORDER BY RefTrafegoID,NProva")
    If RsMatrProva.BOF And RsMatrProva.EOF Then
            MsgBox "Nenhuma prova encontrada para esta matricula.", vbInformation, "CESNet - Aviso!"
            RsMatrProva.Close
            Exit Sub
    End If
    RsMatrProva.Close
    
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
 
    Call cPreview(2, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.FontSize = 14
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("DECLARAÇÃO DE PROVAS EFETUADAS NO ENSINO " & PgNomeEnsino(EnsinoIndice)) / 2)
    
    ObjPreview.Print "DECLARAÇÃO DE PROVAS EFETUADAS NO ENSINO " & PgNomeEnsino(EnsinoIndice)
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontSize = 8
    ObjPreview.Print
    ObjPreview.Print Tab(10); PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc
    ObjPreview.Print
    
    ObjPreview.Print
    ObjPreview.Print Tab(10); "Matrícula: "; Tab(25); PgDadosMatr(MatrID).MatrID
    ObjPreview.Print Tab(10); "Aluno(a): "; Tab(25); PgDadosMatr(MatrID).Nome
    ObjPreview.Print Tab(10); "Nascimento: "; Tab(25); PgDadosMatr(MatrID).Nasc
    ObjPreview.Print Tab(10); "RG: "; Tab(25); PgDadosMatr(MatrID).RG & " (" & PgDadosMatr(MatrID).OE & ")"
    ObjPreview.Print
    ObjPreview.Print Tab(15); "Declaramos para os devidos fins de escolaridade, que o(a) aluno(a) acima efetuou as seguintes provas:"
    ObjPreview.Print
    
    ObjPreview.FontBold = True
    ObjPreview.FontBold = False
        Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoIndice & " ORDER BY DisciplinaID, NProva ASC")
        If RsMatrProva.BOF And RsMatrProva.EOF Then
            Else
                RsMatrProva.MoveFirst
                Do Until RsMatrProva.EOF
                    If DisciplinaID = RsMatrProva.Fields("DisciplinaID") Then
                        Else
                            ObjPreview.FontBold = True
                            ObjPreview.Print
                            ObjPreview.Print Tab(8); "Disciplina: " & PgNomeDisciplina(RsMatrProva.Fields("DisciplinaID"))
                            DisciplinaID = RsMatrProva.Fields("DisciplinaID")
                            ObjPreview.FontBold = False
                    End If
                    ObjPreview.Print Tab(13); RsMatrProva.Fields("DtAvaliacao") & " - " & RsMatrProva.Fields("NProva") & " - " & RsMatrProva.Fields("Assunto")
                    RsMatrProva.MoveNext
                Loop
        End If
        
    
    
    Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
    
End Sub


Public Sub Rpt006(MatrID As String, EnsinoID As Integer, DtDoc As String)
'HISTORICO ESCOLAR
    
    If Trim(DocHistEsc) <> "" Then
        Call RelatorioWord(DocHistEsc, MatrID, PgNomeEnsino(EnsinoID), DtDoc)
        Exit Sub
    End If
    
    Dim RsMatriculaDisciplina   As Recordset
    Dim RsMatriculaEnsino       As Recordset
    Dim RsMatriculaHist         As Recordset
    
    Dim DtConclEnsino           As String
    Dim x, y                    As Integer
    
    Set RsMatriculaDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsMatriculaDisciplina.BOF And RsMatriculaDisciplina.EOF Then
            MsgBox "Esta Matricula não possue DISCIPLINA(S) cadastrada(s)!", vbInformation, "CESNet - Atenção"
            Exit Sub
        Else
            RsMatriculaDisciplina.MoveFirst
    End If
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
    

    Call cPreview(1, UnidadeEnsinoNome, PgDadosUnid(UnidadeEnsino).UA)
    DoEvents
    ObjPreview.Font = CI.Fonte
    ObjPreview.FontBold = True
    ObjPreview.FontItalic = False
    ObjPreview.FontUnderline = False
    ObjPreview.FontSize = 14
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth("HISTÓRICO ESCOLAR - ENSINO " & PgNomeEnsino(EnsinoID)) / 2)
    ObjPreview.CurrentY = 5000 'vertical
    
    
    ObjPreview.Print "HISTÓRICO ESCOLAR - ENSINO " & PgNomeEnsino(EnsinoID)
    ObjPreview.FontBold = CI.Negrito
    ObjPreview.FontSize = CI.tFonte
    
    Dim strNome As String
    'Dim Ensino As String
    
    'Ensino = IIf(PgNomeEnsino(PgMatrEnsino(matrid)) = 0, "NENHUM ENSINO", PgNomeEnsino(PgMatrEnsino(matrid)))
    'Ensino = Ensino & "  " & Mid(String(19, "X"), 1, 20 - Len(Ensino))
    'strNome = Nome '& "  " & Mid(String(49, "X"), 1, 50 - Len(Nome)) & ","
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
      
    ObjPreview.Print Tab(10); "Matrícula: "; Tab(25); PgDadosMatr(MatrID).MatrID
    ObjPreview.Print Tab(10); "Aluno(a): "; Tab(25); PgDadosMatr(MatrID).Nome
    'ObjPreview.Print Tab(10); "Endereço: "; Tab(25); PgDadosMatr(matrid).Endereco
    'ObjPreview.Print Tab(10); "Bairro: "; Tab(25); PgDadosMatr(matrid).Bairro
    'ObjPreview.Print Tab(10); "Município: "; Tab(25); PgDadosMatr(matrid).Munic; " / " & PgDadosMatr(matrid).UF
    'ObjPreview.Print Tab(10); "CEP: "; Tab(25); PgDadosMatr(matrid).CEP
    ObjPreview.Print Tab(10); "Nascimento: "; Tab(25); PgDadosMatr(MatrID).Nasc
    ObjPreview.Print Tab(10); "Filiação: "; Tab(25); PgDadosMatr(MatrID).Pai
    ObjPreview.Print Tab(25); PgDadosMatr(MatrID).Mae
    ObjPreview.Print Tab(10); "Naturalidade: "; Tab(25); PgDadosMatr(MatrID).Natural
    ObjPreview.Print Tab(10); "RG: "; Tab(25); PgDadosMatr(MatrID).RG & " (" & PgDadosMatr(MatrID).OE & ")"
    ObjPreview.Print
    
    'ObjPreview.Print
    
    Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & PgDadosMatr(MatrID).MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
            MsgBox "Erro ao localizar a data de Conclusao do Ensino", vbInformation, "CESNet - Aviso"
            DtConclEnsino = "00/00/0000"
        Else
            RsMatriculaEnsino.MoveFirst
            DtConclEnsino = IIf(IsNull(RsMatriculaEnsino.Fields("DtFinal")), "00/00/0000", RsMatriculaEnsino.Fields("DtFinal"))
            RsMatriculaEnsino.Close
    End If
    ObjPreview.FontBold = True
    ObjPreview.Print Tab(82); "Data de Conclusão do Curso: " & DtConclEnsino
    ObjPreview.FontSize = 8
    
    
    'ObjPreview.Print
    
    ObjPreview.Print
    x = 800 ' horizontalmente
    y = 8500 'verticalmente
    
    ObjPreview.CurrentX = (x / 2) + 50
    ObjPreview.CurrentY = y
    ObjPreview.Print "DISCIPLINA"; Tab(35); "Dt. CONCLUSÃO"; Tab(55); "ESTABELECIMENTO"; Tab(100); "CIDADE/UF"; Tab(131); "SITUAÇÃO FINAL"
    ObjPreview.Line (x / 2, y - 50)-Step(10800, 300), , B
    ObjPreview.FontBold = False
    ' y = y + 250
    
    Do Until RsMatriculaDisciplina.EOF
        y = y + 250
        ObjPreview.CurrentY = y
        ObjPreview.CurrentX = x
        'cor = 6
        
        ObjPreview.Print Tab(7); Trim(PgNomeDisciplina(RsMatriculaDisciplina.Fields("DisciplinaID"))); _
                         Tab(38); IIf(IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) = True, "00/0000", Right(RsMatriculaDisciplina.Fields("DtConclusao"), 7)); _
                         Tab(55); IIf(IsNull(RsMatriculaDisciplina.Fields("Local")), "", Mid(RsMatriculaDisciplina.Fields("Local"), 1, 30)); _
                         Tab(100); IIf(IsNull(RsMatriculaDisciplina.Fields("Cidade")), "", Trim(Mid(RsMatriculaDisciplina.Fields("Cidade"), 1, 30))) & "/" & IIf(IsNull(RsMatriculaDisciplina.Fields("UF")), "", Trim(RsMatriculaDisciplina.Fields("UF"))); _
                         Tab(140); IIf(IsNull(RsMatriculaDisciplina.Fields("DtConclusao")) = True, "INAPTO", "APTO")
                
                'Tab(65); IIf(IsNull(RsMatriculaDisciplina.Fields("Local")), "", RsMatriculaDisciplina.Fields("Local"))

        ObjPreview.Line (x / 2, y + 200)-Step(10800, 200), , B
        'cor = IIf(cor = 6, 15, 6)
        RsMatriculaDisciplina.MoveNext
    Loop
    ObjPreview.FontSize = CI.tFonte
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print Tab(15); "O CENTRO DE ESTUDOS DE JOVENS E ADULTOS - CEJA, utiliza metodologia de ensino modular e individualizado,"
    ObjPreview.Print Tab(10); "sem caráter de série, exigindo-se para aprovação o percentual mínimo de " & Trim(NotaMedia) & "%." 'left(String(5, " "), 5 - Len(Trim(NotaMedia))) & "(" & NotaMedia & "%) determinado pela"
    'ObjPreview.Print Tab(10); "Secretaria Estadual de Educação."
    ObjPreview.Print

    'Call ImpRodape
    'If CI.Preview = False Then
    '    ObjPreview.EndDoc
    'End If
    'If CI.Preview = True Then
    '    Exit Sub
    'End If
    
    'COLOCAR O VERSO DO HISTÓRICO ESCOLAR DIGFITADO NA MATRICULA
    If CI.Preview = False Then
        Set RsMatriculaHist = BD.OpenRecordset("SELECT * FROM HistEscolar WHERE MatrID = '" & MatrID & "' ORDER BY cont")
        If RsMatriculaHist.BOF And RsMatriculaHist.EOF Then
            Else
                RsMatriculaHist.MoveFirst
                If MsgBox("Por favor, vire a folha para imprimir o verso do histórico escolar.", vbInformation + vbYesNo, "CESNet - Aviso") = vbNo Then
                    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc) / 2)
                    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc 'Format(DtDoc, "Long Date")
                    ObjPreview.Print
                    ObjPreview.Print
                    ObjPreview.Print
                    ObjPreview.Print
                    ObjPreview.Print Tab(10); "__________________________________" _
                                    ; Tab(70); "__________________________________"
                    ObjPreview.Print Tab(25); "Secretário(a)" _
                                    ; Tab(85); "Diretor(a)"

                    'Call ImpRodape
                    If CI.Preview = False Then
                        ObjPreview.EndDoc
                    End If
                    Exit Sub
                End If
                If CI.Preview = False Then
                    ObjPreview.EndDoc
                End If
                ObjPreview.Font = CI.Fonte
                ObjPreview.FontBold = True
                ObjPreview.FontItalic = False
                ObjPreview.FontUnderline = False
    
                ObjPreview.Print
                ObjPreview.Print
                ObjPreview.Print
    
                ObjPreview.Print Tab(5); "SÉRIE"; _
                                Tab(17); "ANO"; _
                                Tab(25); "ESTABELECIMENTO DE ENSINO"; _
                                Tab(75); "CIDADE/ESTADO"
    
                ObjPreview.FontBold = False
                ObjPreview.Print
                Do Until RsMatriculaHist.EOF
                    ObjPreview.Print Tab(5); RsMatriculaHist.Fields("Serie"); _
                                    Tab(17); RsMatriculaHist.Fields("Ano"); _
                                    Tab(25); RsMatriculaHist.Fields("Escola"); _
                                    Tab(80); RsMatriculaHist.Fields("Cidade")
                    RsMatriculaHist.MoveNext
                Loop
                ObjPreview.Print
                ObjPreview.Print
                ObjPreview.Print Tab(5); "Este documento é cópia fiel do original arquivado neste estabelecimento de ensino."
        End If
                
    End If
    
    
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    DtDoc = UCase(Trim(Mid(Format(DtDoc, "Long Date"), InStr(Format(DtDoc, "Long Date"), ",") + 1, Len(Format(DtDoc, "Long Date")))))
    ObjPreview.CurrentX = (ObjPreview.ScaleWidth / 2) - (ObjPreview.TextWidth(PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc) / 2)
    ObjPreview.Print PgDadosUnid(UnidadeEnsino).Municipio & ", " & DtDoc
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print
    ObjPreview.Print Tab(10); "__________________________________" _
                     ; Tab(70); "__________________________________"
    ObjPreview.Print Tab(25); "Secretário(a)" _
                     ; Tab(85); "Diretor(a)"
    
    'Call ImpRodape
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If

End Sub


Public Sub ImprCertificado(LadoCert As Integer, nModelo As Integer, MatrID As String, DtCert As String)
   'nModelo: 1= Fundamenta 2=Medio
    'On Error GoTo Trt_Erro
    Dim RsCI                As Recordset
    Dim RsMatrEnsi          As Recordset
    Dim RsMatrDisc          As Recordset
    Dim RsCertificado       As Recordset
    Dim c(99, 2)            As Integer 'c(campo,coordenada) //  coordenada: 1=y 2=x
    Dim campo               As Integer
    Dim LocConclusao        As String
    Dim DtConclusao         As String
    Dim EnsinoID            As Integer
    Dim linha               As Integer
    
    EnsinoID = nModelo
    
    If Form_Impressora.LoadFormCI(True, True, False, True, False, True, True, True, True, True) = False Then
        Exit Sub
    End If
    'Carregar Coordenadas de Impressao
     Set RsCI = BD.OpenRecordset("SELECT * FROM CoordImprCert WHERE Modelo = " & nModelo & " ORDER BY Campo ASC")
    If RsCI.BOF And RsCI.EOF Then
            MsgBox "Erro ao localizar as coordenadas de impreção do certificado." & vbCrLf & _
            "Operação Cancelada!", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsCI.MoveFirst
            Do Until RsCI.EOF
                campo = RsCI.Fields("Campo")
                c(campo, 1) = RsCI.Fields("Me") 'X
                c(campo, 2) = RsCI.Fields("Tp") 'Y
                RsCI.MoveNext
            Loop
    End If
                'Printer.Orientation = 2
            ObjPreview.ScaleMode = 6
            ObjPreview.FontName = CI.Fonte
            
            ObjPreview.FontItalic = CI.Italico
            ObjPreview.FontBold = CI.Negrito
            'ObjPreview.FontUnderline = CI.Sublinhado
            ObjPreview.FontSize = CI.tFonte
    
            ObjPreview.ScaleMode = 6

    
    Select Case LadoCert
        Case 0 'FRENTE
           
            'Dados do Ensino
            Set RsMatrEnsi = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & nModelo)
            If RsMatrEnsi.BOF And RsMatrEnsi.EOF Then
                    MsgBox "Erro ao localizar dados do Ensino", vbInformation, "CESNet - Aviso"
                    RsMatrEnsi.Close
                    Exit Sub
                Else
                    RsMatrEnsi.MoveFirst
                    DtConclusao = cNull(RsMatrEnsi.Fields("DtFinal"))
                    LocConclusao = cNull(RsMatrEnsi.Fields("Local"))
                    RsMatrEnsi.Close
            End If
            
            
            ObjPreview.CurrentX = c(1, 1)
            ObjPreview.CurrentY = c(1, 2)
            ObjPreview.Print Trim(PgDadosUnid("001").NomeCompleto)
    
            ObjPreview.CurrentX = c(2, 1)
            ObjPreview.CurrentY = c(2, 2)
            ObjPreview.Print Trim(PgDadosUnid("001").Endereco) & " - " & Trim(PgDadosUnid("001").Bairro) & " - " & Trim(PgDadosUnid("001").Municipio) & "/" & Trim(PgDadosUnid("001").UF)
    
            ObjPreview.CurrentX = c(3, 1)
            ObjPreview.CurrentY = c(3, 2)
            ObjPreview.Print Trim(PgDadosUnid("001").AtoCriacao)
        
            ObjPreview.CurrentX = c(4, 1)
            ObjPreview.CurrentY = c(4, 2)
            ObjPreview.Print Trim(PgDadosUnid("001").AutorCurso)
    
            ObjPreview.CurrentX = c(5, 1)
            ObjPreview.CurrentY = c(5, 2)
            ObjPreview.Print Trim(PgDadosUnid("001").Nome)
    
            'Nome
            ObjPreview.CurrentX = c(6, 1)
            ObjPreview.CurrentY = c(6, 2)
            ObjPreview.Print PgDadosMatr(MatrID).Nome
    
            'Nacionalidade
            ObjPreview.CurrentX = c(7, 1)
            ObjPreview.CurrentY = c(7, 2)
            ObjPreview.Print PgDadosMatr(MatrID).Nacion
    
            'RG
            ObjPreview.CurrentX = c(8, 1)
            ObjPreview.CurrentY = c(8, 2)
            ObjPreview.Print PgDadosMatr(MatrID).RG
    
            'Orgao emissor
            ObjPreview.CurrentX = c(9, 1)
            ObjPreview.CurrentY = c(9, 2)
            ObjPreview.Print PgDadosMatr(MatrID).OE
    
            'Natural
            ObjPreview.CurrentX = c(10, 1)
            ObjPreview.CurrentY = c(10, 2)
            ObjPreview.Print PgDadosMatr(MatrID).Natural 'left(PgDadosMatr(MatrID).Natural, Len(PgDadosMatr(MatrID).Natural) - 3)
    
            'Unidade Federal
            ObjPreview.CurrentX = c(11, 1)
            ObjPreview.CurrentY = c(11, 2)
            ObjPreview.Print PgDadosMatr(MatrID).NaturalUF ' Right(PgDadosMatr(MatrID).Natural, 2)
    
            'Data Nascimento
            ObjPreview.CurrentX = c(12, 1)
            ObjPreview.CurrentY = c(12, 2)
            ObjPreview.Print left(Trim(PgDadosMatr(MatrID).Nasc), 2)
            ObjPreview.CurrentX = c(13, 1)
            ObjPreview.CurrentY = c(13, 2)
            ObjPreview.Print PgMes(Trim(PgDadosMatr(MatrID).Nasc))
            ObjPreview.CurrentX = c(14, 1)
            ObjPreview.CurrentY = c(14, 2)
            ObjPreview.Print Right(Trim(PgDadosMatr(MatrID).Nasc), 4)
    
            'Nome do Curso (Ensino)
            If nModelo = 1 Then
                ObjPreview.CurrentX = c(18, 1)
                ObjPreview.CurrentY = c(18, 2)
                ObjPreview.Print Trim(PgNomeEnsino(EnsinoID))
            End If
            'Data da Conclusao
            ObjPreview.CurrentX = c(15, 1)
            ObjPreview.CurrentY = c(15, 2)
            ObjPreview.Print left(Trim(DtConclusao), 2)
            ObjPreview.CurrentX = c(16, 1)
            ObjPreview.CurrentY = c(16, 2)
            ObjPreview.Print Trim(PgMes(Trim(DtConclusao)))
            ObjPreview.CurrentX = c(17, 1)
            ObjPreview.CurrentY = c(17, 2)
            ObjPreview.Print Trim(Right(Trim(DtConclusao), 4))
    
            'Municipio da certificação
            ObjPreview.CurrentX = c(19, 1)
            ObjPreview.CurrentY = c(19, 2)
            ObjPreview.Print Trim(PgDadosUnid("001").Municipio)
    
            'Data do Certificado
            ObjPreview.CurrentX = c(20, 1)
            ObjPreview.CurrentY = c(20, 2)
            ObjPreview.Print left(Trim(DtCert), 2)
            ObjPreview.CurrentX = c(21, 1)
            ObjPreview.CurrentY = c(21, 2)
            ObjPreview.Print PgMes(Trim(DtCert))
            ObjPreview.CurrentX = c(22, 1)
            ObjPreview.CurrentY = c(22, 2)
            ObjPreview.Print Right(Trim(DtCert), 4)
   
    
            If CI.Preview = False Then
                    ObjPreview.EndDoc
            End If
                
             
            
            'End If
        
    '***********************************************************************************************
    '************************************  VERSO  **************************************************
    '***********************************************************************************************
        linha = 0
        Case 1 'VERSO

            
    
            Select Case nModelo
'********************************************************************************************************
'********************************************************************************************************
'********************************************************************************************************
                Case 1  'Fundamental
'********************************************************************************************************
'********************************************************************************************************
'********************************************************************************************************

                    'Nome do Aluno
                    ObjPreview.CurrentX = c(23, 1)
                    ObjPreview.CurrentY = c(23, 2) + linha
                    ObjPreview.Print Trim(PgDadosMatr(MatrID).Nome)
            
                    Set RsMatrDisc = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " ORDER BY DtConclusao ASC")
                    If RsMatrDisc.BOF And RsMatrDisc.EOF Then
                        Else
                            RsMatrDisc.MoveFirst
                            Do Until RsMatrDisc.EOF
                                'Disciplina
                                ObjPreview.CurrentX = c(24, 1)
                                ObjPreview.CurrentY = c(24, 2) + linha
                                ObjPreview.Print PgNomeDisciplina(RsMatrDisc.Fields("DisciplinaID"))
                        
                                'Estabelecimento
                                ObjPreview.CurrentX = c(25, 1)
                                ObjPreview.CurrentY = c(25, 2) + linha
                                If pgUsarInstSigla(EnsinoID) = True Then
                                        ObjPreview.Print PgDadosInstEns(RsMatrDisc.Fields("InstID")).Abreviatura
                                    Else
                                        ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("Local")), " ", RsMatrDisc.Fields("Local"))
                                End If
                        
                                'Cidade
                                ObjPreview.CurrentX = c(26, 1)
                                ObjPreview.CurrentY = c(26, 2) + linha
                                If pgUsarCidRed(EnsinoID) = True Then
                                        ObjPreview.Print PgDadosInstEns(RsMatrDisc.Fields("InstID")).CidadeRed 'IIf(IsNull(RsMatrDisc.Fields("CidadeRed")), " ", RsMatrDisc.Fields("CidadeRed"))
                                    Else
                                        ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("Cidade")), " ", RsMatrDisc.Fields("Cidade"))
                                End If
                                'Estado
                                ObjPreview.CurrentX = c(27, 1)
                                ObjPreview.CurrentY = c(27, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("UF")), " ", RsMatrDisc.Fields("UF"))
                        
                                'Data da Conclusao
                                ObjPreview.CurrentX = c(28, 1)
                                ObjPreview.CurrentY = c(28, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("DtConclusao")), "00", Mid(RsMatrDisc.Fields("DtConclusao"), 4, 2))
                                ObjPreview.CurrentX = c(29, 1)
                                ObjPreview.CurrentY = c(29, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("DtConclusao")), "0000", Right(RsMatrDisc.Fields("DtConclusao"), 4))
                            
                                'Mensao
                                ObjPreview.CurrentX = c(30, 1)
                                ObjPreview.CurrentY = c(30, 2) + linha
                                ObjPreview.Print MensaoHB
                                
                                linha = linha + ObjPreview.TextHeight(MensaoHB) '+ 30
                                
                                RsMatrDisc.MoveNext
                            Loop
                    End If

                    'Data do Certificado
                    ObjPreview.CurrentX = c(31, 1)
                    ObjPreview.CurrentY = c(31, 2) + linha
                    ObjPreview.Print left(Trim(DtCert), 2)
                    ObjPreview.CurrentX = c(32, 1)
                    ObjPreview.CurrentY = c(32, 2) + linha
                    ObjPreview.Print PgMes(Trim(DtCert))
                    ObjPreview.CurrentX = c(33, 1)
                    ObjPreview.CurrentY = c(33, 2) + linha
                    ObjPreview.Print Right(Trim(DtCert), 4)
    
            
'********************************************************************************************************
'********************************************************************************************************
'********************************************************************************************************
                Case 2  'Medio
'********************************************************************************************************
'********************************************************************************************************
'********************************************************************************************************
        
            
                    Set RsMatrDisc = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " ORDER BY DtConclusao ASC")
                    If RsMatrDisc.BOF And RsMatrDisc.EOF Then
                        Else
                            RsMatrDisc.MoveFirst
                            Do Until RsMatrDisc.EOF
                                'Disciplina
                                ObjPreview.CurrentX = c(23, 1)
                                ObjPreview.CurrentY = c(23, 2) + linha
                                ObjPreview.Print PgNomeDisciplina(RsMatrDisc.Fields("DisciplinaID"))
                        
                                'Forma de Estudo
                                ObjPreview.CurrentX = c(24, 1)
                                ObjPreview.CurrentY = c(24, 2) + linha
                                ObjPreview.Print FormEstudo 'PgNomeDisciplina(RsMatrDisc.Fields("DisciplinaID"))
                        
                                'Data
                                ObjPreview.CurrentX = c(25, 1)
                                ObjPreview.CurrentY = c(25, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("DtConclusao")), "00/0000", Mid(RsMatrDisc.Fields("DtConclusao"), 4, 8))
                        
                            
                        
                                'Total de Horas (NAO TEM)
                        
                                'Estabelecimento
                                ObjPreview.CurrentX = c(27, 1)
                                ObjPreview.CurrentY = c(27, 2) + linha
                                If pgUsarInstSigla(EnsinoID) = True Then
                                        ObjPreview.Print PgDadosInstEns(IIf(IsNull(RsMatrDisc.Fields("InstID")), "0", RsMatrDisc.Fields("InstID"))).Abreviatura
                                    Else
                                        ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("Local")), " ", RsMatrDisc.Fields("Local"))
                                End If

                                'If IsNull(RsMatrDisc.Fields("Abrev")) Then
                                '        If IsNull(RsMatrDisc.Fields("InstID")) Then
                                '                ObjPreview.Print " "
                                '            Else
                                '
                                '        End If
                                '    Else
                                '        ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("Abrev")), " ", RsMatrDisc.Fields("Abrev"))
                                'End If
                        
                        
                                'Cidade
                                ObjPreview.CurrentX = c(28, 1)
                                ObjPreview.CurrentY = c(28, 2) + linha
                                'ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("Cidade")), " ", RsMatrDisc.Fields("Cidade"))
                                If pgUsarCidRed(EnsinoID) = True Then
                                        ObjPreview.Print PgDadosInstEns(IIf(IsNull(RsMatrDisc.Fields("InstID")), 0, RsMatrDisc.Fields("InstID"))).CidadeRed 'IIf(IsNull(RsMatrDisc.Fields("CidadeRed")), " ", RsMatrDisc.Fields("CidadeRed"))
                                    Else
                                        ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("Cidade")), " ", RsMatrDisc.Fields("Cidade"))
                                End If
                            
                                'Estado
                                ObjPreview.CurrentX = c(29, 1)
                                ObjPreview.CurrentY = c(29, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("UF")), " ", RsMatrDisc.Fields("UF"))
                            
                        
                                'Data da Conclusao MES
                                ObjPreview.CurrentX = c(30, 1)
                                ObjPreview.CurrentY = c(30, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("DtConclusao")), "00", Mid(RsMatrDisc.Fields("DtConclusao"), 4, 2))
                                'Debug.Print IIf(IsNull(RsMatrDisc.Fields("DtConclusao")), "00", Mid(RsMatrDisc.Fields("DtConclusao"), 4, 2))
                                'Data da Conclusao ANO
                                ObjPreview.CurrentX = c(31, 1)
                                ObjPreview.CurrentY = c(31, 2) + linha
                                ObjPreview.Print IIf(IsNull(RsMatrDisc.Fields("DtConclusao")), "0000", Right(RsMatrDisc.Fields("DtConclusao"), 4))
                            
                                'Mensao
                                ObjPreview.CurrentX = c(32, 1)
                                ObjPreview.CurrentY = c(32, 2) + linha
                                ObjPreview.Print MensaoHB
                                
                                linha = linha + ObjPreview.TextHeight(MensaoHB) ' + 30
                                
                                RsMatrDisc.MoveNext
                            Loop
                    End If
                    
                    Set RsCertificado = BD.OpenRecordset("SELECT * FROM Certificado WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
                    If RsCertificado.BOF And RsCertificado.EOF Then
                            RsCertificado.Close
                        Else
                            RsCertificado.MoveFirst
                            'Curso Anterior
                            ObjPreview.CurrentX = c(33, 1)
                            ObjPreview.CurrentY = c(33, 2)
                            ObjPreview.Print IIf(IsNull(RsCertificado.Fields("CursoAnt")), " ", RsCertificado.Fields("CursoAnt"))
            
                        
                            'Estabelecimento
                            ObjPreview.CurrentX = c(34, 1)
                            ObjPreview.CurrentY = c(34, 2)
                            ObjPreview.Print IIf(IsNull(RsCertificado.Fields("Estabelecimento")), " ", RsCertificado.Fields("Estabelecimento"))
                
                            'Localidade e UF
                            ObjPreview.CurrentX = c(35, 1)
                            ObjPreview.CurrentY = c(35, 2)
                            ObjPreview.Print IIf(IsNull(RsCertificado.Fields("LocalUF")), " ", RsCertificado.Fields("LocalUF"))
                    
            
                            'Outras Hab
                            ObjPreview.CurrentX = c(36, 1)
                            ObjPreview.CurrentY = c(36, 2)
                            ObjPreview.Print IIf(IsNull(RsCertificado.Fields("OutrasHab")), " ", RsCertificado.Fields("OutrasHab"))
            
                            'Obs
                            ObjPreview.CurrentX = c(37, 1)
                            ObjPreview.CurrentY = c(37, 2)
                            ObjPreview.Print IIf(IsNull(RsCertificado.Fields("Obs")), " ", RsCertificado.Fields("Obs"))
                    End If
                    'REGISTRO
                    ObjPreview.CurrentX = c(38, 1)
                    ObjPreview.CurrentY = c(38, 2) + linha
                    ObjPreview.Print IIf(IsNull(RsCertificado.Fields("DOReg")), " ", RsCertificado.Fields("DOReg"))
                
                    'Folha
                    ObjPreview.CurrentX = c(39, 1)
                    ObjPreview.CurrentY = c(39, 2) + linha
                    ObjPreview.Print IIf(IsNull(RsCertificado.Fields("DOFolha")), " ", RsCertificado.Fields("DOFolha"))
            
                    'Livro
                    ObjPreview.CurrentX = c(40, 1)
                    ObjPreview.CurrentY = c(40, 2) + linha
                    ObjPreview.Print IIf(IsNull(RsCertificado.Fields("DOLivro")), " ", RsCertificado.Fields("DOLivro"))
                
                    'Data da Listagem
                    ObjPreview.CurrentX = c(41, 1)    '42
                    ObjPreview.CurrentY = c(41, 2) + linha '42
                    ObjPreview.Print IIf(IsNull(RsCertificado.Fields("DODtPublicacao")), " ", RsCertificado.Fields("DODtPublicacao"))
            
                    'fOLHA
                    ObjPreview.CurrentX = c(43, 1)
                    ObjPreview.CurrentY = c(43, 2) + linha
                    ObjPreview.Print IIf(IsNull(RsCertificado.Fields("DOFolhaDO")), " ", RsCertificado.Fields("DOFolhaDO"))


                    'Local
                    ObjPreview.CurrentX = c(44, 1)
                    ObjPreview.CurrentY = c(44, 2)
                    ObjPreview.Print PgDadosUnid("001").Municipio
            
                    'Data do Certificado
                    ObjPreview.CurrentX = c(45, 1)
                    ObjPreview.CurrentY = c(45, 2)
                    ObjPreview.Print DtCert 'Left(Trim(DtCert), 2)
                    'ObjPreview.CurrentX = c(46, 1) '32
                    'ObjPreview.CurrentY = c(46, 2) + Linha
                    'ObjPreview.Print PgMes(Trim(DtCert))
                    'ObjPreview.CurrentX = c(47, 1) '33
                    'ObjPreview.CurrentY = c(47, 2) + Linha
                    'ObjPreview.Print Right(Trim(DtCert), 4)
    

            End Select
    End Select
    If CI.Preview = False Then
        ObjPreview.EndDoc
    End If
    Exit Sub
Trt_Erro:
    MsgBox Err.Description, vbInformation, Err.Number
    Resume Next
End Sub
