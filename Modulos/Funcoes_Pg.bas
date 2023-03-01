Attribute VB_Name = "Funcoes_Pg"
Option Explicit
Dim RsEnsino            As Recordset
Dim RsDisciplina        As Recordset
Dim RsSerie             As Recordset
Dim RsModulo            As Recordset
Dim RsMatricula         As Recordset
Dim RsDeficiencia       As Recordset
Dim RsAviso             As Recordset
Dim RsGrupoUsu          As Recordset
Type DadosMatricula
    MatrID      As String
    DtMatr      As String
    DtRetorno   As String
    UnidMatr    As String
    UnidMatrID  As String
    Nome        As String
    Endereco    As String
    Numero      As String
    Compl       As String
    Bairro      As String
    Munic       As String
    UF          As String
    CEP         As String
    Sexo        As String
    Nasc        As String
    Mail        As String
    Cel         As String
    Tel1        As String
    Tel2        As String
    RG          As String
    OE          As String
    CPF         As String
    CertNasc    As String
    Natural     As String
    NaturalUF   As String
    EstCivil    As String
    Nacion      As String
    Pai         As String
    Mae         As String
    DtNascMae   As String
    DtNascPai   As String
    Raca        As String
    OpcaoRel    As String
    TpSang      As String
    Deficiencia As Integer
    ValCard     As Date
    NumAnt      As String
    NumCenso    As String
    NumConexao  As String
    Obs         As String
End Type

Type DadosUnidade
    UnidadeID       As String
    Nome            As String
    NomeCompleto    As String
    Endereco        As String
    Bairro          As String
    Municipio       As String
    UF              As String
    CEP             As String
    Cnpj            As String
    AtoCriacao      As String
    AutorCurso      As String
    CodEscolar      As String
    UA              As String
End Type

Type DadosInstEnsino
    InstID          As Integer
    Nome            As String
    Abreviatura     As String
    Cidade          As String
    CidadeRed       As String
    UF              As String
End Type

Type DadosMatriculaCurso
    DtInicio        As String
    DtFinal         As String
    Local           As String
End Type

Public Versao               As String
Public SoftwareID           As String 'Deve ser analisado para na hora da inst.
                                      'gerar um num de 13 dig.

Public SisNota              As Boolean ' onde: True = sistema de notas por percentual
                                       '       False = sistema de notas em aprovado e nao aprovado
Public NotaMedia            As String
Public MaxAcessos           As Long
'Public pgNumMinDiscipl        As Integer 'Numero min de discipl nao obrigatorias
Public DeslWin              As Boolean

Public UnidadeEnsino        As String 'Numero da Unidade Local
Public UnidadeEnsinoNome    As String 'Nome da Uniade Local

Public VincModulo           As Boolean 'Vincular provas ao emp modulo
Public MaxDisciplCursando   As Integer
Public CartConjugada        As Integer
Public InatDias             As Integer
Public BloquearInativo      As Boolean 'Bloquear ou nao os alunos inativos
Public RPTime               As Integer 'Resposta de Provas Tempo de atualizacao de tela
Public RPMax                As Integer 'Resposta de Provas Maximo de Resp para apresentar
Public RPMaxReg             As Integer 'Resposta de Provas Maximo de Registros pesq
Public NumRepro             As Integer
Public TextoRepro           As String
Public MensaoHB             As String
Public MensaoNH             As String
Public FormEstudo           As String
Public ultManutencao        As String
Public DocHistEsc           As String
Public BloqRenovVencida     As Integer 'Bloquear se o aluno nao renovou a matr

 

Public Function PgDadosCurso(MatrID As String, CursoID As Integer) As DadosMatriculaCurso
    Dim sSQL    As String
    Dim Rst     As Recordset
    
    sSQL = "SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & CursoID
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            MsgBox "Erro ao localizar dados Curso", vbInformation, "Aviso"
        Else
            Rst.MoveFirst
            PgDadosCurso.DtInicio = IIf(IsNull(Rst.Fields("DtInicio")), "", Rst.Fields("DtInicio"))
            PgDadosCurso.DtFinal = IIf(IsNull(Rst.Fields("DtFinal")), "", Rst.Fields("DtFinal"))
            PgDadosCurso.Local = IIf(IsNull(Rst.Fields("Local")), "", Rst.Fields("Local"))
    End If
    Rst.Close
End Function

Public Function PgMatrEnsino(Matr As String, Optional Concluido = False)
    'Concluido = True ==> Ensino concluido
    'Concluido = False ==> Ensino nao concluiu
    Dim RsMatriculaEnsino As Recordset
    If Concluido = False Then
            'Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & Matr & "' AND DtInicio <> Null AND IsNull(DtFinal)") ' = " & Null) 'IsNull(DtFinal)")
            Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & Matr & "' AND Trancado = FALSE AND IsNull(DtFinal)") ' = " & Null)
        Else
            Set RsMatriculaEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & Matr & "' AND Trancado = FALSE AND DtFinal <> Null")
    End If
    If RsMatriculaEnsino.BOF And RsMatriculaEnsino.EOF Then
            PgMatrEnsino = 0
        Else
            PgMatrEnsino = RsMatriculaEnsino.Fields("EnsinoID")
    End If
    RsMatriculaEnsino.Close
End Function

Public Function PgNomeEnsino(EnsinoID As Integer)
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino WHERE ID = " & EnsinoID)
    If RsEnsino.BOF And RsEnsino.EOF Then
            PgNomeEnsino = 0
        Else
            RsEnsino.MoveFirst
            PgNomeEnsino = RsEnsino.Fields("Descr")
    End If
    RsEnsino.Close
End Function
Public Function PgIDEnsino(Ensino As String)
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino WHERE Descr = '" & Ensino & "'")
    If RsEnsino.BOF And RsEnsino.EOF Then
            PgIDEnsino = 0
        Else
            RsEnsino.MoveFirst
            PgIDEnsino = RsEnsino.Fields("ID")
    End If
    RsEnsino.Close
End Function
Public Function PgIDDisciplina(Disciplina As String)
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina WHERE Descr = '" & Disciplina & "'")
    If RsDisciplina.BOF And RsDisciplina.EOF Then
            PgIDDisciplina = 0
        Else
            RsDisciplina.MoveFirst
            PgIDDisciplina = RsDisciplina.Fields("ID")
    End If
    RsDisciplina.Close
End Function
Public Function PgNomeDisciplina(DisciplinaID As Integer)
    Set RsDisciplina = BD.OpenRecordset("SELECT * FROM Disciplina WHERE ID = " & DisciplinaID)
    If RsDisciplina.BOF And RsDisciplina.EOF Then
            PgNomeDisciplina = 0
        Else
            RsDisciplina.MoveFirst
            PgNomeDisciplina = RsDisciplina.Fields("Descr")
    End If
    RsDisciplina.Close
End Function
Public Function PgNomeDef(DefID As Integer)
    Set RsDeficiencia = BD.OpenRecordset("SELECT * FROM Deficiencias WHERE ID = " & DefID)
    If RsDeficiencia.BOF And RsDeficiencia.EOF Then
            PgNomeDef = " "
        Else
            RsDeficiencia.MoveFirst
            PgNomeDef = RsDeficiencia.Fields("Descr")
    End If
    RsDeficiencia.Close
End Function
Public Function PgIDDef(Deficiencia As String)
    Set RsDeficiencia = BD.OpenRecordset("SELECT * FROM Deficiencias WHERE Descr = '" & Deficiencia & "'")
    If RsDeficiencia.BOF And RsDeficiencia.EOF Then
            PgIDDef = 0
        Else
            RsDeficiencia.MoveFirst
            PgIDDef = RsDeficiencia.Fields("ID")
    End If
    RsDeficiencia.Close
End Function

Public Function PgIDModulo(Modulo As String)
    Set RsModulo = BD.OpenRecordset("SELECT * FROM Modulo WHERE Descr = '" & Modulo & "'")
    If RsModulo.BOF And RsModulo.EOF Then
            PgIDModulo = 0
        Else
            RsModulo.MoveFirst
            PgIDModulo = RsModulo.Fields("ID")
    End If
    RsModulo.Close
End Function
Public Function PgNomeModulo(ModuloID As Integer)
    Set RsModulo = BD.OpenRecordset("SELECT * FROM Modulo WHERE ID = " & ModuloID)
    If RsModulo.BOF And RsModulo.EOF Then
            PgNomeModulo = 0
        Else
            RsModulo.MoveFirst
            PgNomeModulo = RsModulo.Fields("Descr")
    End If
    RsModulo.Close
End Function

Public Function PgNomeSerie(SerieID As Integer)

    Set RsSerie = BD.OpenRecordset("SELECT * FROM Serie WHERE ID = " & SerieID)
    If RsSerie.BOF And RsSerie.EOF Then
            PgNomeSerie = 0
        Else
            RsSerie.MoveFirst
            PgNomeSerie = RsSerie.Fields("Descr")
    End If
    RsSerie.Close
End Function
Public Function PgIDSerie(Serie As String)
    Set RsSerie = BD.OpenRecordset("SELECT * FROM Serie WHERE Descr = '" & Serie & "'")
    If RsSerie.BOF And RsSerie.EOF Then
            PgIDSerie = 0
        Else
            RsSerie.MoveFirst
            PgIDSerie = RsSerie.Fields("ID")
    End If
    RsSerie.Close
End Function

Public Function PgDadosMatr(Matricula As String) As DadosMatricula
    On Error GoTo TrtErroMatr
    Dim Matr As DadosMatricula
    
    Set RsMatricula = BD.OpenRecordset("SELECT * FROM Matriculas WHERE MatrID = '" & Matricula & "'")
    If RsMatricula.BOF And RsMatricula.EOF Then
            'Matr = ""
            PgDadosMatr = Matr
        Else
            RsMatricula.MoveFirst
            With RsMatricula
                Matr.MatrID = .Fields("MatrID")
                Matr.DtMatr = Format(.Fields("DtMat"), "dd/mm/yyyy")
                Matr.DtRetorno = IIf(IsNull(.Fields("DtRetorno")), "", .Fields("DtRetorno"))
                Matr.Nome = .Fields("Nome")
                If IsNull(.Fields("UnidadeID")) Then
                        Matr.UnidMatrID = "" 'IIf(IsNull(.Fields("UnidadeID")), " ", left("000", 3 - Len(.Fields("UnidadeID"))) & .Fields("UnidadeID"))
                    Else
                        Matr.UnidMatrID = left("000", 3 - Len(.Fields("UnidadeID"))) & .Fields("UnidadeID")
                End If
                Matr.UnidMatr = PgDadosUnid(IIf(IsNull(.Fields("UnidadeID")), 0, .Fields("UnidadeID"))).Nome
                Matr.Endereco = IIf(IsNull(.Fields("End")), " ", .Fields("End"))
                Matr.Numero = IIf(IsNull(.Fields("Numero")), " ", .Fields("Numero"))
                Matr.Compl = IIf(IsNull(.Fields("Compl")), " ", .Fields("Compl"))
                Matr.Bairro = IIf(IsNull(.Fields("Bai")), " ", .Fields("Bai"))
                Matr.Munic = IIf(IsNull(.Fields("Mun")), " ", .Fields("Mun"))
                Matr.UF = IIf(IsNull(.Fields("UF")), " ", .Fields("UF"))
                Matr.CEP = IIf(IsNull(.Fields("CEP")), " ", .Fields("CEP"))
                Matr.Sexo = IIf(IsNull(.Fields("Sexo")), " ", .Fields("Sexo"))
        
                Matr.Nasc = IIf(IsNull(.Fields("Nasc")), "", .Fields("Nasc"))
        
                Matr.Mail = IIf(IsNull(.Fields("Mail")), " ", .Fields("Mail"))
                Matr.Cel = IIf(IsNull(.Fields("CEL")), " ", .Fields("CEL"))
                Matr.Tel1 = IIf(IsNull(.Fields("Tel1")), " ", .Fields("Tel1"))
                Matr.Tel2 = IIf(IsNull(.Fields("Tel2")), " ", .Fields("Tel2"))
                Matr.RG = IIf(IsNull(.Fields("RG")), " ", .Fields("RG"))
                Matr.OE = IIf(IsNull(.Fields("OE")), " ", .Fields("OE"))
                Matr.CPF = IIf(IsNull(.Fields("CPF")), " ", .Fields("CPF"))
                Matr.CertNasc = IIf(IsNull(.Fields("CertNasc")), " ", .Fields("CertNasc"))
                Matr.Natural = IIf(IsNull(.Fields("Natural")), " ", .Fields("Natural"))
                Matr.NaturalUF = IIf(IsNull(.Fields("NaturalUF")), " ", .Fields("NaturalUF"))
                Matr.EstCivil = IIf(IsNull(.Fields("EstCivil")), " ", .Fields("EstCivil"))
                Matr.Nacion = IIf(IsNull(.Fields("Nacion")), " ", .Fields("Nacion"))
                
                Matr.Mae = IIf(IsNull(.Fields("Mae")), " ", .Fields("Mae"))
                Matr.DtNascMae = IIf(IsNull(.Fields("DtNascMae")), "", .Fields("DtNascMae"))
                Matr.Pai = IIf(IsNull(.Fields("Pai")), " ", .Fields("Pai"))
                Matr.DtNascPai = IIf(IsNull(.Fields("DtNascPai")), "", .Fields("DtNascPai"))
                
                Matr.Raca = IIf(IsNull(.Fields("Raca")), "", .Fields("Raca"))
                Matr.OpcaoRel = IIf(IsNull(.Fields("OpcaoRel")), "", .Fields("OpcaoRel"))
                Matr.TpSang = IIf(IsNull(.Fields("TpSang")), "", .Fields("TpSang"))
                
                Matr.ValCard = IIf(IsNull(.Fields("ValCard")), Date, .Fields("ValCard"))
                Matr.NumAnt = IIf(IsNull(.Fields("NumAnt")), "", .Fields("NumAnt"))
                
                Matr.NumCenso = IIf(IsNull(.Fields("NumCenso")), "", .Fields("NumCenso"))
                Matr.NumConexao = IIf(IsNull(.Fields("NumConexao")), "", .Fields("NumConexao"))
                
                Matr.Deficiencia = IIf(IsNull(.Fields("DefID")), "0", .Fields("DefID"))
                
                Matr.Obs = IIf(IsNull(.Fields("Obs")), " ", .Fields("Obs"))
        End With
        PgDadosMatr = Matr
    End If
    RsMatricula.Close
    Exit Function
TrtErroMatr:
    MsgBox "Modulo: PgDadosMatr" & vbCrLf & "Descrição: " & Err.Description, vbCritical, "CESNet - Erro nº" & Err.Number
    Resume Next
End Function

Public Function PgDadosUnid(Optional UnidadeID As String) As DadosUnidade
    Dim RsUnidade As Recordset
    UnidadeID = IIf(Trim(UnidadeID) = "", UnidadeEnsino, UnidadeID)
    Dim Unid        As DadosUnidade
    
    Set RsUnidade = BD.OpenRecordset("SELECT * FROM Unidades WHERE UnidID = '" & left("000", 3 - Len(Trim(UnidadeID))) & UnidadeID & "'")
    If RsUnidade.BOF And RsUnidade.EOF Then
            'Matr = ""
            PgDadosUnid = Unid
        Else
            RsUnidade.MoveFirst
            With RsUnidade
                
                Unid.UnidadeID = .Fields("UnidID")
                Unid.Nome = .Fields("Nome")
                Unid.NomeCompleto = IIf(IsNull(.Fields("NomeCompleto")), " ", .Fields("NomeCompleto"))
                Unid.Endereco = IIf(IsNull(.Fields("End")), " ", .Fields("End"))
                Unid.Bairro = IIf(IsNull(.Fields("Bai")), " ", .Fields("Bai"))
                Unid.Municipio = IIf(IsNull(.Fields("Mun")), " ", .Fields("Mun"))
                Unid.UF = IIf(IsNull(.Fields("UF")), " ", .Fields("UF"))
                Unid.CEP = IIf(IsNull(.Fields("CEP")), " ", .Fields("CEP"))
                Unid.Cnpj = IIf(IsNull(.Fields("CNPJ")), " ", .Fields("CNPJ"))
                Unid.AtoCriacao = IIf(IsNull(.Fields("Criacao")), " ", .Fields("Criacao"))
                Unid.AutorCurso = IIf(IsNull(.Fields("AutoCurso")), " ", .Fields("AutoCurso"))
                Unid.CodEscolar = IIf(IsNull(.Fields("CodEscolar")), " ", .Fields("CodEscolar"))
                Unid.UA = IIf(IsNull(.Fields("UA")), " ", .Fields("UA"))
'               Txt_Criacao.Text = PgDadosUnid(UnidadeEnsino).AtoCriacao
'               Txt_AutoCurso.Text = PgDadosUnid(UnidadeEnsino).AutorCurso

        
                'Unid.Nasc = IIf(IsNull(.Fields("Nasc")), "", .Fields("Nasc"))
        
                'Unid.Mail = IIf(IsNull(.Fields("Mail")), " ", .Fields("Mail"))
                'Unid.Cel = IIf(IsNull(.Fields("CEL")), " ", .Fields("CEL"))
                'Unid.Tel1 = IIf(IsNull(.Fields("Tel1")), " ", .Fields("Tel1"))
                'Unid.Tel2 = IIf(IsNull(.Fields("Tel2")), " ", .Fields("Tel2"))
                'Unid.RG = IIf(IsNull(.Fields("RG")), " ", .Fields("RG"))
                'Unid.OE = IIf(IsNull(.Fields("OE")), " ", .Fields("OE"))
                'Unid.CertNasc = IIf(IsNull(.Fields("CertNasc")), " ", .Fields("CertNasc"))
                'Unid.Natural = IIf(IsNull(.Fields("Natural")), " ", .Fields("Natural"))
                'Unid.EstCivil = IIf(IsNull(.Fields("EstCivil")), " ", .Fields("EstCivil"))
                'Unid.Nacion = IIf(IsNull(.Fields("Nacion")), " ", .Fields("Nacion"))
                'Unid.Mae = IIf(IsNull(.Fields("Mae")), " ", .Fields("Mae"))
                'Unid.Pai = IIf(IsNull(.Fields("Pai")), " ", .Fields("Pai"))
                'Unid.ValCard = IIf(IsNull(.Fields("ValCard")), Date, .Fields("ValCard"))
                'Unid.NumAnt = IIf(IsNull(.Fields("NumAnt")), "", .Fields("NumAnt"))
                'Unid.Deficiencia = IIf(IsNull(.Fields("DefID")), "0", .Fields("DefID"))
                'Unid.Obs = IIf(IsNull(.Fields("Obs")), " ", .Fields("Obs"))
        End With
        PgDadosUnid = Unid
    End If
    RsUnidade.Close
End Function
Public Function PgDadosInstEns(Optional InstEnsID As Integer) As DadosInstEnsino
    'InstEnsID = IIf(Trim(InstEnsID) = "", UnidadeEnsino, InstEnsID)
    Dim RsInstEnsino    As Recordset
    Dim InstEns         As DadosInstEnsino
    
    Set RsInstEnsino = BD.OpenRecordset("SELECT * FROM InstEnsino WHERE ID = " & InstEnsID)
    If RsInstEnsino.BOF And RsInstEnsino.EOF Then
            'Matr = ""
            PgDadosInstEns = InstEns
        Else
            RsInstEnsino.MoveFirst
            With RsInstEnsino
                
                'InstEns.InstEnsID = .Fields("InstID")
                InstEns.Nome = .Fields("Descr")
                InstEns.Abreviatura = IIf(IsNull(.Fields("Sigla")), " ", .Fields("Sigla"))
                InstEns.Cidade = IIf(IsNull(.Fields("Cidade")), " ", .Fields("Cidade"))
                InstEns.CidadeRed = IIf(IsNull(.Fields("CidadeRed")), " ", .Fields("CidadeRed"))
                InstEns.UF = IIf(IsNull(.Fields("UF")), " ", .Fields("UF"))
        End With
        PgDadosInstEns = InstEns
    End If
    RsInstEnsino.Close
End Function
Public Function PgRespUsu(UsuID As Integer)
    Dim RsTMP As Recordset
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuID)
    If RsTMP.BOF And RsTMP.EOF Then
            PgRespUsu = ""
        Else
            RsTMP.MoveFirst
            PgRespUsu = RsTMP.Fields("Responsavel")
    End If
    RsTMP.Close
End Function
Public Sub PgRegrasSis()
    'On Error GoTo TrtErro
    Dim VerSis      As String
    Dim a, b, c     As String
    Dim RsRegras    As Recordset
    Dim RsUnidade   As Recordset
Voltar:
    Set RsRegras = BD.OpenRecordset("SELECT * FROM Config")
    RsRegras.MoveFirst
    
    'Validar a versao do sistema
    Versao = IIf(IsNull(RsRegras.Fields("Versao")), "", RsRegras.Fields("Versao"))
    a = App.Major
    b = App.Minor
    c = App.Revision
    VerSis = Mid("00", 1, 2 - Len(a)) & a & "." & Mid("00", 1, 2 - Len(b)) & b & "." & Mid("000", 1, 3 - Len(c)) & c
    Select Case Versao
        Case Is > VerSis
            'If MsgBox("Versão do CESNet invalida, por favor atualize!!", vbInformation + vbYesNo, "CESNet - Aviso") = vbYes Then
            '        AtualizarVersao
            '        End
            '    Else
            '        End
            'End If
            
        Case Is < VerSis
            RsRegras.Edit
            RsRegras.Fields("Versao") = VerSis
            RsRegras.Update
            RsRegras.Close
            Form_SQLExecuteAuto.CarregarFormulario
            Versao = VerSis
            
            GoTo Voltar
        Case Is = VerSis
        Case Else
            RsRegras.Edit
            RsRegras.Fields("Versao") = VerSis
            RsRegras.Update
            Versao = VerSis
            'Call Form_ChkTabsBD.IniChkBD
    End Select
    
    
    
    UnidadeEnsino = Mid(String(3, "0"), 1, 3 - Len(RsRegras.Fields("Unidade"))) & RsRegras.Fields("Unidade")
    MaxAcessos = RsRegras.Fields("MaxAcessos")
    NotaMedia = RsRegras.Fields("NotaMedia")
    'pgNumMinDiscipl = IIf(IsNull(RsRegras.Fields("pgNumMinDiscipl")), "0", RsRegras.Fields("pgNumMinDiscipl"))
    UnidadeEnsino = Mid(String(3, "0"), 1, 3 - Len(RsRegras.Fields("Unidade"))) & RsRegras.Fields("Unidade")
    
    BloquearInativo = RsRegras.Fields("BloquearInativo")
    'SoftwareID = NumSerieInst 'String(5, "0")
    
    DocHistEsc = IIf(IsNull(RsRegras.Fields("nmDocHistEsc")), "", RsRegras.Fields("nmDocHistEsc"))
    
    
    BloqRenovVencida = IIf(IsNull(RsRegras.Fields("BloqRenovVencida")), 0, RsRegras.Fields("BloqRenovVencida"))
    
    
    VincModulo = RsRegras.Fields("VincModulos")
    '***************************************************************************************************
    'NA PROXIMA VERSAO SEPARAR O QUANTITATIVO DE EMPR DE MODULOS DO QUANT. DE DISCIPLINAS
    'CURSADAS SIMULTANIAMENTE
    MaxDisciplCursando = RsRegras.Fields("MaxModulosEmpr")
    'MsgBox "Olhar Regrasis"
    '*******************************************************************************************************
    
    SisNota = RsRegras.Fields("SisNota")
    
    RPTime = IIf(IIf(IsNull(RsRegras.Fields("RPTime")), 0, RsRegras.Fields("RPTime")) < 1, 1, RsRegras.Fields("RPTime"))
    RPMax = IIf(IIf(IsNull(RsRegras.Fields("RPMax")), 0, RsRegras.Fields("RPMax")) < 1, 1, RsRegras.Fields("RPMax"))
    RPMaxReg = IIf(IIf(IsNull(RsRegras.Fields("RPMaxReg")), 0, RsRegras.Fields("RPMaxReg")) < 1, 1, RsRegras.Fields("RPMaxReg"))
    
    NumRepro = IIf(IsNull(RsRegras.Fields("NumRepro")), 0, Trim(RsRegras.Fields("NumRepro")))
    InatDias = IIf(IsNull(RsRegras.Fields("InatDias")), 5, Trim(RsRegras.Fields("InatDias")))
    CartConjugada = IIf(IsNull(RsRegras.Fields("CartConjugada")), 0, Trim(RsRegras.Fields("CartConjugada")))
    
    
    DeslWin = RsRegras.Fields("DeslWin")
    
    Set RsUnidade = BD.OpenRecordset("SELECT * FROM Unidades WHERE UnidID = '" & UnidadeEnsino & "'")
    
    If RsUnidade.BOF And RsUnidade.EOF Then
            UnidadeEnsinoNome = "< NENHUMA UNIDADE DE ENSINO CADASTRADA. >"
        Else
            RsUnidade.MoveFirst
            UnidadeEnsinoNome = RsUnidade.Fields("Nome")
    End If
    
    ultManutencao = IIf(IsNull(RsRegras.Fields("UltManutencao")), "", Trim(RsRegras.Fields("UltManutencao")))
    If IsNull(RsRegras.Fields("UltManutencao")) Then
        Else
            If RsRegras.Fields("UltManutencao") < Date Then
                If Status_Serv = 1 Or Status_Serv = 2 Then
                    MsgBox "Procedimentos de Manutenção a base de dados não executada. Favor avisar ao Administrador do CESNet.", vbInformation, "CESNet - Aviso"
                End If
            End If
    End If
    
    MensaoHB = IIf(IsNull(RsRegras.Fields("HB")), "", Trim(RsRegras.Fields("HB")))
    MensaoNH = IIf(IsNull(RsRegras.Fields("NH")), "", Trim(RsRegras.Fields("NH")))
    FormEstudo = IIf(IsNull(RsRegras.Fields("FormEstudo")), "", Trim(RsRegras.Fields("FormEstudo")))
    
    RsUnidade.Close
    RsRegras.Close
    Exit Sub
TrtErro:
    Call RegLogErros(Err.Number, Err.Description, "PgRegrasSis", UsuarioID)
    MsgBox "Modulo: PgRegrasSis" & Chr(13) & "- Descrição: " & Err.Description, vbCritical, "Erro n.: " & Err.Number
    Resume Next
    
End Sub




Public Function PgNomeGrupoUsu(GrupoID As Integer) As String
    Set RsGrupoUsu = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE GrupoID = " & GrupoID)
    If RsGrupoUsu.BOF And RsGrupoUsu.EOF Then
            PgNomeGrupoUsu = ""
        Else
            RsGrupoUsu.MoveFirst
            PgNomeGrupoUsu = Trim(RsGrupoUsu.Fields("Nome"))
    End If
    RsGrupoUsu.Close
End Function

Public Function PgIDGrupoUsu(Grupo As String) As Integer
    Set RsGrupoUsu = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE Nome = " & Grupo)
    If RsGrupoUsu.BOF And RsGrupoUsu.EOF Then
            PgIDGrupoUsu = 0
        Else
            RsGrupoUsu.MoveFirst
            PgIDGrupoUsu = Trim(RsGrupoUsu.Fields("GrupoID"))
    End If
    RsGrupoUsu.Close
End Function
Public Function PgNomeUF(Sigla As String)
      Select Case Sigla

            Case "AC"
                PgNomeUF = "ACRE"
            Case "AL"
                PgNomeUF = "ALAGOAS"
            Case "AM"
                PgNomeUF = "AMAZONAS"
            Case "AP"
                PgNomeUF = "AMAPA"
            Case "BA"
                PgNomeUF = "BAHIA"
            Case "CE"
                PgNomeUF = "CEARA"
            Case "DF"
                PgNomeUF = "DISTRITO FEDERAL"
            Case "ES"
                PgNomeUF = "ESPIRITO SANTO"
            Case "GO"
                PgNomeUF = "GOIAS"
            Case "MA"
                PgNomeUF = "MARANHAO"
            Case "MG"
                PgNomeUF = "MINAS GERAIS"
            Case "MS"
                PgNomeUF = "MATO GROSSO DO SUL"
            Case "MT"
                PgNomeUF = "MATO GROSSO"
            Case "PA"
                PgNomeUF = "PARA"
            Case "PB"
                PgNomeUF = "PARAIBA"
            Case "PE"
                PgNomeUF = "PERNAMBUCO"
            Case "PI"
                PgNomeUF = "PIAUI"
            Case "PR"
                PgNomeUF = "PARANA"
            Case "RJ"
                PgNomeUF = "RIO DE JANEIRO"
            Case "RN"
                PgNomeUF = "RIO GRANDE SO NORTE"
            Case "RO"
                PgNomeUF = "RONDONIA"
            Case "RS"
                PgNomeUF = "RIO GRANDE DO SUL"
            Case "SC"
                PgNomeUF = "SANTA CATARINA"
            Case "SE"
                PgNomeUF = "SERGIPE"
            Case "SP"
                PgNomeUF = "SAO PAULO"
            Case "TO"
                PgNomeUF = "TOCANTINS"
            Case Else
                PgNomeUF = "Sigla não cadastrada no CESNet"
        End Select
End Function
Public Function PgAviso(MatrID As String) As Boolean
    Dim Avisar      As Boolean
    Dim Bloquear    As Boolean
    'true - bloquear aluno
    'false - nao bloquear aluno
    Set RsAviso = BD.OpenRecordset("SELECT * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "'")
    If RsAviso.BOF And RsAviso.EOF Then
            PgAviso = False
            Exit Function
        Else
            RsAviso.MoveFirst
            Avisar = False
            Bloquear = False
            PgAviso = False
    End If
    Do Until RsAviso.EOF
        If RsAviso.Fields("DtAvisar") <= Date Then
            Avisar = True
        End If
        If RsAviso.Fields("DtBloqueio") <= Date Then
            Bloquear = True
        End If
        RsAviso.MoveNext
    Loop
            
    PgAviso = Bloquear
            
    If Avisar = True Then
            Call Form_MatriculaAvisoPreview.CarregarForm(MatrID, Bloquear) '(MatrID, RsAviso.Fields("DtInclusao"), _
                                                        IIf(IsNull(RsAviso.Fields("DtBloqueio")), "", RsAviso.Fields("DtBloqueio")), _
                                                        RsAviso.Fields("Texto"), PgRespUsu(RsAviso.Fields("UsuID")), PgAviso)
        Else
            MsgBox "Matricula com restrição." & Chr(13) & "Por favor, procure a direção!", vbInformation, "CESNet - Aviso!"
    End If
            
    
End Function
Public Function PgProfDisciplina(ProfID As Integer, DisciplID As Integer) As Boolean
'Informa: True - Professor da disciplina
'         False - Nao é prof da disciplina
    Dim RsProfessorDisciplina As Recordset
    Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina WHERE ProfID = " & ProfID & " AND DisciplinaID = " & DisciplID)
    If RsProfessorDisciplina.BOF And RsProfessorDisciplina.EOF Then
            PgProfDisciplina = False
        Else
            PgProfDisciplina = True
    End If
    RsProfessorDisciplina.Close
End Function
Public Function PgNomeGrupo(GrupoID As String) As String
    Dim RsUsuGrupo As Recordset
    Set RsUsuGrupo = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE GrupoID = " & GrupoID)
    If RsUsuGrupo.BOF And RsUsuGrupo.EOF Then
            PgNomeGrupo = " "
        Else
            PgNomeGrupo = RsUsuGrupo.Fields("Nome")
    End If
    RsUsuGrupo.Close
End Function
Public Function PgIDGrupo(NGrupo As String) As String
    Dim RsUsuGrupo As Recordset
    Set RsUsuGrupo = BD.OpenRecordset("SELECT * FROM UsuGrupo WHERE Nome = '" & NGrupo & "'")
    If RsUsuGrupo.BOF And RsUsuGrupo.EOF Then
            PgIDGrupo = " "
        Else
            PgIDGrupo = RsUsuGrupo.Fields("GrupoID")
    End If
    RsUsuGrupo.Close
End Function
Public Function PgIDOcorrenciaConclusao(Ocorrencia As String) As String
    Dim RsOcorrenciaConclusao As Recordset
    Set RsOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM OcorrenciaConclusao WHERE Descr = '" & Ocorrencia & "'")
    If RsOcorrenciaConclusao.BOF And RsOcorrenciaConclusao.EOF Then
            PgIDOcorrenciaConclusao = " "
        Else
            PgIDOcorrenciaConclusao = RsOcorrenciaConclusao.Fields("OcorrenciaID")
    End If
    RsOcorrenciaConclusao.Close
End Function
Public Function PgNomeOcorrenciaConclusao(OcorrenciaID As String) As String
    Dim RsOcorrenciaConclusao As Recordset
    Set RsOcorrenciaConclusao = BD.OpenRecordset("SELECT * FROM OcorrenciaConclusao WHERE OcorrenciaID = " & OcorrenciaID)
    If RsOcorrenciaConclusao.BOF And RsOcorrenciaConclusao.EOF Then
            PgNomeOcorrenciaConclusao = " "
        Else
            PgNomeOcorrenciaConclusao = RsOcorrenciaConclusao.Fields("Descr")
    End If
    RsOcorrenciaConclusao.Close
End Function

Public Sub PgSoftwareID()
    On Error GoTo TrtErro
    Dim Arquivo As String
    Dim i As Integer
    Dim sf(2) As String
    Dim linha As String
    Arquivo = FreeFile
    
    Open App.path & "\CESNetID.dat" For Input As Arquivo
    
    For i = 1 To 1
        Line Input #Arquivo, linha
        sf(i) = Crypto(linha)
    Next
    
    Close #Arquivo
    SoftwareID = sf(1)
    Exit Sub
'TRATANDO DOS ERROS
TrtErro:
    If Err.Number = 53 Then
        Call RegLogErros(Err.Number, Err.Description, "PgSoftwareID", UsuarioID)
        Arquivo = FreeFile
        Open App.path & "\CESNetid.dat" For Output As Arquivo
        Print #Arquivo, Crypto(NumSerieInst)
        Close #Arquivo
    End If
    End
End Sub
Public Function PgMes(Dt As String) As String
    Dt = Mid(Dt, 4, 2)
    Select Case Dt
        Case "01"
            PgMes = "JANEIRO"
        Case "02"
            PgMes = "FEVEREIRO"
        Case "03"
            PgMes = "MARÇO"
        Case "04"
            PgMes = "ABRIL"
        Case "05"
            PgMes = "MAIO"
        Case "06"
            PgMes = "JUNHO"
        Case "07"
            PgMes = "JULHO"
        Case "08"
            PgMes = "AGOSTO"
        Case "09"
            PgMes = "SETEMBRO"
        Case "10"
            PgMes = "OUTUBRO"
        Case "11"
            PgMes = "NOVEMBRO"
        Case "12"
            PgMes = "DEZEMBRO"
    End Select
End Function

Public Function PgStatusMatricula(MatrID As String) As String
    On Error GoTo TratErro
    Dim Status              As String
'Informa se a matricula acima esta ATIVA ou INATIVA
    Dim Intervalo           As String
    Dim RsMatrProva         As Recordset
    Dim RsMatrOrientacao    As Recordset
    Dim RsMatrEnsino        As Recordset
    '************* CHECA MATRICULA
    Intervalo = DateDiff("d", PgDadosMatr(MatrID).DtMatr, Date)
    Intervalo = IIf(Val(Intervalo) <= 0, 0, Intervalo)
    If Intervalo <= InatDias Then
            PgStatusMatricula = "ATIVO"
            Exit Function
        Else
            PgStatusMatricula = "INATIVO"
    End If
    'RETORNO
    If Trim(PgDadosMatr(MatrID).DtRetorno) = "" Then
            PgStatusMatricula = "INATIVO"
        Else
            Intervalo = DateDiff("d", PgDadosMatr(MatrID).DtRetorno, Date)
            Intervalo = IIf(Val(Intervalo) <= 0, 0, Intervalo)
            If Intervalo <= InatDias Then
                    PgStatusMatricula = "ATIVO"
                    Exit Function
                Else
                    PgStatusMatricula = "INATIVO"
            End If
        
    End If
    
    '************ CHECAR PROVAS
    Set RsMatrProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & PgMatrEnsino(MatrID) & " ORDER BY DtAvaliacao")
    If RsMatrProva.BOF And RsMatrProva.EOF Then
            PgStatusMatricula = "INATIVO"
        Else
            RsMatrProva.MoveLast
            Intervalo = DateDiff("d", RsMatrProva.Fields("DtAvaliacao"), Date)
            Intervalo = IIf(Val(Intervalo) < 0, 0, Intervalo)
            If Intervalo <= InatDias Then
                    PgStatusMatricula = "ATIVO"
                    RsMatrProva.Close
                    Exit Function
                Else
                    PgStatusMatricula = "INATIVO"
            End If
    End If
    RsMatrProva.Close
    '***************CHECA ORIENTACAO
    Set RsMatrOrientacao = BD.OpenRecordset("SELECT * FROM MatriculaOrientacao WHERE MatrID = '" & MatrID & "' ORDER BY DtOrientacao")
    If RsMatrOrientacao.BOF And RsMatrOrientacao.EOF Then
            PgStatusMatricula = "INATIVO"
        Else
            If IsNull(RsMatrOrientacao.Fields("DtOrientacao")) Then
                    PgStatusMatricula = "INATIVO"
                    'RsMatrOrientacao.Close
                Else
                    Intervalo = DateDiff("d", RsMatrOrientacao.Fields("DtOrientacao"), Date)
                    Intervalo = IIf(Val(Intervalo) < 0, 0, Intervalo)
                    If Intervalo <= InatDias Then
                            PgStatusMatricula = "ATIVO"
                            RsMatrOrientacao.Close
                            Exit Function
                        Else
                            PgStatusMatricula = "INATIVO"
                    End If
            End If
    End If
    RsMatrOrientacao.Close
    'PgStatusMatricula = Status
    
    
    
        '***** CHECA SE O ALUNO ESTA INATIVO E BLOQUEA-LO ********
    Dim RsAviso As Recordset
    If PgStatusMatricula = "INATIVO" And BloquearInativo = True Then
        Set RsAviso = BD.OpenRecordset("SELECT * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "' AND Codigo=1")
        If RsAviso.BOF And RsAviso.EOF Then
            RsAviso.AddNew
            RsAviso.Fields("MatrID") = MatrID
            RsAviso.Fields("DtInclusao") = Date
            RsAviso.Fields("Codigo") = "1"
            RsAviso.Fields("Texto") = "<MATRICULA INATIVA - Aviso automatico do CESNet.>"
            RsAviso.Fields("Avisar") = True
            RsAviso.Fields("DtAvisar") = Date
            RsAviso.Fields("Bloquear") = BloquearInativo
            RsAviso.Fields("DtBloqueio") = IIf(BloquearInativo = True, Date, Null)
            RsAviso.Fields("UsuID") = UsuarioID
            RsAviso.Fields("DtHr") = Now
            RsAviso.Update
            RsAviso.Close
        End If

    End If
    
    '********************************************
    
    '########################################################################################
    '###  CHECA SE O ALUNO ESTA COM A RENOVACAO DE  MATRICULA VENCIDA E BLOQUEIA-O
    Dim stBloqMatr As Boolean
    
    stBloqMatr = False
    Set RsMatrEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & PgMatrEnsino(MatrID))
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
        Else
            If Date > CDate(IIf(IsNull(RsMatrEnsino.Fields("dtRenovacao")), Date, RsMatrEnsino.Fields("dtRenovacao"))) Then
                stBloqMatr = True
            End If
    End If
    RsMatrEnsino.Close
    '########################################################################################
    If stBloqMatr = True And BloqRenovVencida <> 0 Then
        Set RsAviso = BD.OpenRecordset("SELECT * FROM MatriculaAviso WHERE MatrID = '" & MatrID & "' AND Codigo=2")
        If RsAviso.BOF And RsAviso.EOF Then
            RsAviso.AddNew
            RsAviso.Fields("MatrID") = MatrID
            RsAviso.Fields("DtInclusao") = Date
            RsAviso.Fields("Codigo") = "2"
            RsAviso.Fields("Texto") = "<MATRICULA NÃO RENOVADA - Aviso automatico do CESNet.>"
            RsAviso.Fields("Avisar") = True
            RsAviso.Fields("DtAvisar") = Date
            RsAviso.Fields("Bloquear") = BloquearInativo
            RsAviso.Fields("DtBloqueio") = IIf(BloquearInativo = True, Date, Null)
            RsAviso.Fields("UsuID") = UsuarioID
            RsAviso.Fields("DtHr") = Now
            RsAviso.Update
            RsAviso.Close
        End If

    End If
    
    
    
    Exit Function
TratErro:
    PgStatusMatricula = "ERRO"
End Function
Public Function PgNomeProfessor(ProfID As Integer) As String
    Dim RsProfessor As Recordset
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores WHERE ProfID = " & ProfID)
    If RsProfessor.BOF And RsProfessor.EOF Then
            'MsgBox "Professor não cadastrado na base de dados.", vbInformation, "CESNet - Aviso!"
            PgNomeProfessor = ""
        Else
            RsProfessor.MoveFirst
            PgNomeProfessor = RsProfessor.Fields("Nome")
            'Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina WHERE Chv = '" & Usuario & "'")
    End If
    RsProfessor.Close
End Function
Public Function PgIDProfessor(ProfNome As String) As Integer
    Dim RsProfessor As Recordset
    Set RsProfessor = BD.OpenRecordset("SELECT * FROM Professores WHERE Nome = '" & ProfNome & "'")
    If RsProfessor.BOF And RsProfessor.EOF Then
            'MsgBox "Professor não cadastrado na base de dados.", vbInformation, "CESNet - Aviso!"
            PgIDProfessor = 0
        Else
            RsProfessor.MoveFirst
            PgIDProfessor = RsProfessor.Fields("ProfID")
            'Set RsProfessorDisciplina = BD.OpenRecordset("SELECT * FROM ProfessorDisciplina WHERE Chv = '" & Usuario & "'")
    End If
    RsProfessor.Close
End Function
Public Function PgNomeDoc(strDescr As String) As String
    'Pega o nome do documento da declaracao
    Dim RsTMP As Recordset
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Documentos WHERE Descr = '" & strDescr & "'")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
            PgNomeDoc = ""
        Else
            RsTMP.MoveFirst
            PgNomeDoc = RsTMP.Fields("nArq")
            RsTMP.Close
    End If
End Function

Public Function PgDescArqDoc(strArquivo As String) As String
    'Pega a desccricao do nome do arquivo para Declaracao
    Dim RsTMP As Recordset
    
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Documentos WHERE nArq = '" & strArquivo & "'")
    If RsTMP.BOF And RsTMP.EOF Then
            RsTMP.Close
            PgDescArqDoc = ""
        Else
            RsTMP.MoveFirst
            PgDescArqDoc = RsTMP.Fields("Descr")
            RsTMP.Close
    End If
End Function
Public Function PgCorFundo() As String
    If Trim(UsuarioID) = "" Then
        PgCorFundo = vbWhite
        Exit Function
    End If
    Dim RsTMP As Recordset
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Usuario WHERE UsuarioID = " & UsuarioID)
    If RsTMP.BOF And RsTMP.EOF Then
            PgCorFundo = vbWhite
            RsTMP.Close
        Else
            RsTMP.MoveFirst
            PgCorFundo = IIf(IsNull(RsTMP.Fields("CorFundo")), vbWhite, RsTMP.Fields("CorFundo"))
            RsTMP.Close
    End If
        
End Function
Public Function pgNumMinDiscipl(intCurso As Integer) As Integer
    Dim RsEnsino As Recordset
    
    Set RsEnsino = BD.OpenRecordset("SELECT * FROM Ensino WHERE ID = " & intCurso)
    If RsEnsino.BOF And RsEnsino.EOF Then
            MsgBox "Erro ao localizar Ensino na tabEnsino"
            Exit Function
        Else
            RsEnsino.MoveFirst
            pgNumMinDiscipl = IIf(IsNull(RsEnsino.Fields("NumMinDiscipl")), 0, RsEnsino.Fields("NumMinDiscipl"))
    End If
    RsEnsino.Close
    
    
End Function
Public Function pgUsarCidRed(EnsID As Integer) As Boolean
    'True - Usar
    'False - Nao Usar
    Dim Rst As Recordset
    Set Rst = BD.OpenRecordset("SELECT * FROM Ensino WHERE ID = " & EnsID) 'ORDER BY ID")
    If Rst.BOF And Rst.EOF Then
            pgUsarCidRed = False
        Else
            Rst.MoveFirst
            If Rst.Fields("UsarCidReduzida") = 1 Then
                    pgUsarCidRed = True
                Else
                    pgUsarCidRed = False
            End If
    End If
    Rst.Close
End Function
Public Function pgUsarInstSigla(EnsID As Integer) As Boolean
    'True - Usar
    'False - Nao Usar
    
    Dim Rst As Recordset
    Set Rst = BD.OpenRecordset("SELECT * FROM Ensino WHERE ID = " & EnsID) 'ORDER BY ID")
    If Rst.BOF And Rst.EOF Then
            pgUsarInstSigla = False
        Else
            Rst.MoveFirst
            If Rst.Fields("UsarInstSigla") = 1 Then
                    pgUsarInstSigla = True
                Else
                    pgUsarInstSigla = False
            End If
            
    End If
    
    Rst.Close
End Function

