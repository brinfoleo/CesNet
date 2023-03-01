Attribute VB_Name = "Nucleo_ConclusaoEnsino"
Option Explicit

'Versão 1.00 do Nucleo de Conclusão de Ensino
'O Objetivo deste modulo é checar e concluir series, disciplinas e ensinos.

Public Sub NormalizarEnsino(MatrID As String, EnsinoID As Integer)
    If Trim(MatrID) = "" Or EnsinoID = 0 Then Exit Sub
    Dim RsMatrEnsino    As Recordset
    Dim RsMatrDiscipl   As Recordset
    Dim RsMatrSerie     As Recordset
    
    'Verificar se o ensino solicitado esta cadastrado
    Set RsMatrEnsino = BD.OpenRecordset("SELECT * FROM MatriculaEnsino WHERE MatrId = '" & MatrID & "' AND EnsinoID = " & EnsinoID)
    If RsMatrEnsino.BOF And RsMatrEnsino.EOF Then
            MsgBox "Curso não Cadastrado para esta matrícula.", vbInformation, "CESNet - Aviso"
            RsMatrEnsino.Close
        Else
            RsMatrEnsino.MoveLast
            'Checa ambiguidade nos ensinos e exclui um deles
            If RsMatrEnsino.RecordCount > 1 Then
                If MsgBox("Exite ambiguidade de ensino nesta matricula." & vbCrLf & _
                          "Deseja EXCLUIR uma delas?", vbInformation + vbYesNo, "CESNet - Aviso") = vbYes Then
                          BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & EnsinoID & " AND IsNull(DtInicio) AND IsNull(DtFinal)"

                          RsMatrEnsino.MoveFirst
                    Else
                        MsgBox "Impossivel continuar com dois cursos iguais." & vbCrLf & "Operacao cancelada", vbInformation, "CESNet - Aviso"
                        RsMatrEnsino.Close
                        Exit Sub
                End If
            End If
            
    End If
End Sub
Public Function EliminarDuplicidadeEnsino()
    On Error Resume Next
    
    Dim Rst     As Recordset
    Dim sSQL    As String
    Dim tReg    As Long
    Dim rCont   As Integer
    
    
    
    'Pega os dados do Ensino
    'Set Rst = New ADODB.Recordset
    sSQL = "SELECT * FROM MatriculaEnsino ORDER BY MatrID"
    
    Set Rst = BD.OpenRecordset(sSQL) ', BD, adOpenDynamic
    If Rst.BOF And Rst.EOF Then
            tReg = 0
            'Rodape (tReg)
        Else
            RegLog "0", String(120, "=")
            tReg = Rst.RecordCount
            Rst.MoveFirst
            Do Until Rst.EOF
                rCont = NumCursosCad(Rst.Fields("MatrID"), Rst.Fields("EnsinoID"))
                If Rst.Fields("EnsinoID") = 0 Then
                    
                    RegLog "Manutencao", "Matr.:" & Rst.Fields("MatrID") & _
                                     " EnsinoID: 0" & _
                                     " Dt.Inicio:" & IIf(IsNull(Rst.Fields("DtInicio")), "  /  /    ", Rst.Fields("DtInicio")) & _
                                     " Dt.Final:" & IIf(IsNull(Rst.Fields("DtFinal")), "  /  /    ", Rst.Fields("DtFinal")) & _
                                     " Local:" & IIf(IsNull(Rst.Fields("Local")), "", Rst.Fields("Local"))
                    BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & Rst.Fields("MatrID") & "' AND EnsinoID=0"
                End If
                
                If rCont >= 2 Then
                    BD.Execute "DELETE * FROM MatriculaEnsino WHERE MatrID = '" & Rst.Fields("MatrID") & "' AND EnsinoID=" & Rst.Fields("EnsinoID") & " AND " & _
                               "DtFinal IS NULL AND Local IS NULL"
                End If
                Rst.MoveNext
                'Status "Processado...", tReg
            Loop
            RegLog "", String(120, "=")
    End If
    'BD.Close
    
    MsgBox "Fim do processo!", vbInformation, "Aviso"
    
End Function
Private Function NumCursosCad(MatrID As String, Curso As Integer) As Integer
    Dim Rst     As Recordset
    Dim sSQL    As String
    
    sSQL = "SELECT * FROM MatriculaEnsino WHERE MatrID = '" & MatrID & "' AND EnsinoID = " & Curso
    'Rst.Open sSQL, BD
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            NumCursosCad = 0
        Else
            NumCursosCad = Rst.RecordCount
            If NumCursosCad >= 2 Then
                Rst.MoveFirst
                Do Until Rst.EOF
                    RegLog "Manutencao", NumCursosCad & " - Matr.:" & Rst.Fields("MatrID") & _
                                     " EnsinoID:" & Rst.Fields("EnsinoID") & _
                                     " Dt.Inicio:" & IIf(IsNull(Rst.Fields("DtInicio")), "  /  /    ", Rst.Fields("DtInicio")) & _
                                     " Dt.Final:" & IIf(IsNull(Rst.Fields("DtFinal")), "  /  /    ", Rst.Fields("DtFinal")) & _
                                     " Local:" & IIf(IsNull(Rst.Fields("Local")), "", Rst.Fields("Local"))
                    Rst.MoveNext
                Loop
                 
            End If
    End If
    Rst.Close
End Function


