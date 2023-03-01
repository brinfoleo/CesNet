Attribute VB_Name = "Funcoes_Chk_Grv"
Option Explicit
Public Function ChecarProvas(MatrID As String, EnsinoID As Integer, DisciplinaID As Integer, SerieID As Integer)
'CHECA SE EXISTEM MAIS PROVAS PARA O ALUNO
'True - Concluido
'False - Não Concluido
    Dim RsTrafego           As Recordset
    Dim RsProva             As Recordset
    Dim RsMatriculaProva    As Recordset
    Dim RefTrafegoID        As Integer
    Dim nProva              As String
    'Checa as provas de acordo com a tab.trafego
    Set RsTrafego = BD.OpenRecordset("SELECT * FROM Trafego WHERE " & _
                    "EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " AND SerieID = " & SerieID)
    If RsTrafego.BOF And RsTrafego.EOF Then
            MsgBox "Erro ao localizar trafego", vbInformation, "CESNet - Atenção"
            ChecarProvas = False
            Exit Function
        Else
            RsTrafego.MoveFirst
    End If
    Do Until RsTrafego.EOF
        RefTrafegoID = RsTrafego.Fields("RefTrafegoID")
        Set RsProva = BD.OpenRecordset("SELECT * FROM Provas WHERE RefTrafegoID = " & RefTrafegoID & " ORDER BY NPROVA")
        RsProva.MoveFirst
        Do Until RsProva.EOF
            nProva = RsProva.Fields("NProva")
            'Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND RefTrafegoID = " & RefTrafegoID & " and NProva = '" & nProva & "'")
            Set RsMatriculaProva = BD.OpenRecordset("SELECT * FROM MatriculaProva WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " and NProva = '" & nProva & "'")
            If RsMatriculaProva.BOF And RsMatriculaProva.EOF Then
                    ChecarProvas = False
                    Exit Function
                Else
                    If RsMatriculaProva.Fields("Aprovado") = True Then
                            ChecarProvas = True
                        Else
                            ChecarProvas = False
                            Exit Function
                    End If
            End If
            RsProva.MoveNext
        Loop
        RsTrafego.MoveNext
    Loop
End Function

Public Function Chk_ConcDisciplina(MatrID As String, EnsinoID As Integer, DisciplinaID As Integer) As Boolean
    'Checa todas as series concluidas e se a disc ja pode ser concluida
    Dim RsMatriculaSerie As Recordset
    Set RsMatriculaSerie = BD.OpenRecordset("SELECT * FROM MatriculaSerie WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & DisciplinaID & " ORDER BY SerieID") ' AND SerieID = " & SerieID)
    If RsMatriculaSerie.BOF And RsMatriculaSerie.EOF Then
            MsgBox "Erro no acesso ao Banco de Dados. Por favor chame o Suporte", vbInformation, "Atenção"
        Else
            RsMatriculaSerie.MoveFirst
            Do Until RsMatriculaSerie.EOF
                ''Debug.Print RsMatriculaSerie.Fields("SerieID")
                If RsMatriculaSerie.Fields("Aprovado") = False Then
                        Chk_ConcDisciplina = False
                        Exit Function
                    Else
                        Chk_ConcDisciplina = True
                End If
                RsMatriculaSerie.MoveNext
            Loop
    End If
End Function
Public Function Chk_ConcEnsino(MatrID As String, EnsinoID As Integer) As Boolean
    'True - concluiu o ensino
    'False - nao concluiu o ensino
    Dim RsGrdEnsDiscipl     As Recordset  'Tabela GradeEnsinoDisciplinas
    Dim RsMatrDisciplina    As Recordset
    Dim DiscNOConcl         As Integer 'Disciplinas nao obrigatorias concluidas
    DiscNOConcl = 0
    Set RsGrdEnsDiscipl = BD.OpenRecordset("SELECT * FROM GradeEnsinoDisciplinas WHERE EnsinoID = " & EnsinoID & " ORDER BY DisciplinaID")
    If RsGrdEnsDiscipl.BOF And RsGrdEnsDiscipl.EOF Then
            MsgBox "Erro ao localizar as Disciplinas refente ao Ensino!", vbInformation, "CESNet - Aviso"
            Chk_ConcEnsino = False
            Exit Function
        Else
            RsGrdEnsDiscipl.MoveFirst
    End If
    Do Until RsGrdEnsDiscipl.EOF
        'Checa as Disciplinas
        Set RsMatrDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND DisciplinaID = " & RsGrdEnsDiscipl.Fields("DisciplinaID") & " AND DtConclusao <> NULL")
        If RsMatrDisciplina.BOF And RsMatrDisciplina.EOF Then
                'Nao achou a Disciplina concluida
                Chk_ConcEnsino = False
                If RsGrdEnsDiscipl.Fields("Obrigatoria") = True Then
                        Exit Function
                    Else
                        RsGrdEnsDiscipl.MoveNext
                End If
            Else
                'Achou a Disciplina concluida
                If RsGrdEnsDiscipl.Fields("Obrigatoria") = False Then
                    DiscNOConcl = DiscNOConcl + 1
                End If
                RsGrdEnsDiscipl.MoveNext
                
        End If
    Loop
    'MsgBox "ss"
    If DiscNOConcl >= pgNumMinDiscipl(EnsinoID) Then
            Chk_ConcEnsino = True
        Else
            'Caso o ensino nao possua disciplinas nao obrigatorias o ensino deve ser concluido
            If DiscNOConcl = 0 Then
                    Chk_ConcEnsino = True
                Else
                    Chk_ConcEnsino = False
            End If
    End If
    
End Function
Public Sub Grv_ConcEnsino(MatrID As String, EnsinoID As Integer)
    'Pega a ultima data da conclusao da disciplina e grava como
    'data de conclusao do ensino
    
    Dim RsMatrEnsino As Recordset
    Dim RsMatrDisciplina As Recordset
    Set RsMatrDisciplina = BD.OpenRecordset("SELECT * FROM MatriculaDisciplina WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " ORDER BY DtConclusao")
    If RsMatrDisciplina.BOF And RsMatrDisciplina.EOF Then
            MsgBox "Erro ao localizar as disciplinas concluidas", vbInformation, "CESNet - Aviso"
            RsMatrDisciplina.Close
            Exit Sub
        Else
            RsMatrDisciplina.MoveLast
            Set RsMatrEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
            If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
                    MsgBox "Erro ao localizar o ensino", vbInformation, "CESNet - Aviso"
                    RsMatrEnsino.Close
                    RsMatrDisciplina.Close
                    Exit Sub
                Else
                    RsMatrEnsino.Edit
                    RsMatrEnsino.Fields("DtFinal") = RsMatrDisciplina.Fields("DtConclusao")
                    RsMatrEnsino.Fields("Local") = RsMatrDisciplina.Fields("Local")
                    RsMatrEnsino.Update
            End If
    End If
    RsMatrEnsino.Close
    RsMatrDisciplina.Close
End Sub
Public Function ChkExisteArquivo(ByVal sFileName As String) As Boolean


'// Check if File Exists

  Dim sFile As String

  On Error Resume Next

  ChkExisteArquivo = False

  sFile = Dir$(sFileName)
  If (Len(sFile) > 0) And (Err = 0) Then
      ChkExisteArquivo = True
  End If

End Function
Public Function ChkDtProxRenovacao(sMatr As String, DtMatricula As String) As String
    'CDate(.TextMatrix(.Rows - 1, 1)) + 212
    '#######################################################
    '### 15/09/2011
    '### Retorna a data da Proxima Renovacao
    '### Periodo em acordo com Decreto de 2011
    '#######################################################
    Dim sSQL    As String
    Dim Rst     As Recordset
    Dim DtRet   As String
    If IsDate(DtMatricula) = False Then
        ChkDtProxRenovacao = "00/00/0000"
        Exit Function
    End If
    sSQL = "SELECT * FROM MatriculaRetorno WHERE MatrID = '" & sMatr & "' AND TpMov = 2 ORDER BY DtRetorno"
    
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            DtRet = CDate(DtMatricula) + 212
        Else
            Rst.MoveLast
            DtRet = CDate(Rst.Fields("DtRetorno")) + 212
    End If
    Rst.Close
    ChkDtProxRenovacao = DtRet
    
    'BD.Execute "UPDATE MatriculaEnsino SET DtRenovacao=" & DtRet & " WHERE MatrID = '" & sMatr & "' AND DtInicio = " & DtMatricula
End Function


