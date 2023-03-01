Attribute VB_Name = "NumSerie"
Option Explicit
Dim N(3, 99)    As String
Dim i           As Integer
Dim ix          As Integer
'================================================================
'== DADOS PARA CALCULO DE CHA DE ATIVACAO                      ==
'==                                                            ==
'== Data Base: 01/12/2005                                      ==
'==                                                            ==
'================================================================

Public Function NumSerieInst() As String

    Dim DtInstall1 As String
    Dim DtResult1 As String

    Dim NSerieDV1 As String
    Dim NInst1 As String
    Dim Grupo1 As String
    Dim DtBase As Date
    Dim NSerieHD As String
    NSerieHD = IIf(NumSerieHD < 0, NumSerieHD * -1, NumSerieHD)
    
    DtBase = "01/12/2005"
    'Calcula o Digito verificador do num do HD
    NSerieDV1 = CalcDV(NSerieHD)
    
    DtInstall1 = Date 'Data da Instalacao
    If CDate(DtInstall1) < CDate(DtBase) Then
        MsgBox "Data do sistema Invalida por favor verifique", vbInformation, "atençao"
        Exit Function
    End If
    DtResult1 = Trim(Val(Format(DtInstall1, "yyyymmdd")) - Val(Format(DtBase, "yyyymmdd")))
    
    If Len(DtResult1) > 6 Then
        MsgBox "Data de instalação invalida"
        Exit Function
    End If
    DtResult1 = Mid(String(6, "0"), 1, 6 - Len(DtResult1)) & DtResult1
    NInst1 = DtResult1 & NSerieHD & NSerieDV1
    'Text2.Text = DtResult1 & NSerieHD & NSerieDV1
    'O grupo sera selecionado de acor do com o numero do dv do HD
    Grupo1 = NGrupo(NSerieDV1)
    'Text4.Text = Grupo1
    NumSerieInst = ConvNum(Grupo1, NInst1) & NSerieDV1
    'Text6.Text = Text3.Text 'chave de instalacao


End Function
Private Function NumSerieHD() As String
On Error GoTo TrtErro
    Dim Unidade As String
    Dim lSerial As String 'Long
    Dim fso As New FileSystemObject, drvDrive As Drive
    'Unidade = Left(CurDir, 2)
    Unidade = left(App.path, 2)
    Set drvDrive = fso.GetDrive(left(fso.GetDriveName(Unidade), 2))
    lSerial = drvDrive.SerialNumber
    NumSerieHD = lSerial
    Exit Function
TrtErro:
    NumSerieHD = left(Date, 2) & Right(Date, 4) & "00" & Mid(Date, 3, 2)
End Function

Private Function CalcDV(Num As String)
    'Dim i As Integer
    Dim fator As Integer
    Dim Result As Integer
    
    fator = 1
    
    For i = 1 To Len(Num)
        Result = Result + IIf(Len(fator * (Mid(Num, i, 1))) > 2, Right(fator * (Mid(Num, i, 1)), 1), fator * (Mid(Num, i, 1)))
        fator = IIf(fator = 1, 2, 1)
    Next
    CalcDV = Right(Result, 1)
End Function
Private Function ConvLetra(Grupo As String, Letra As String) As String
    Dim npeg As String
    Dim texto As String
    Select Case Grupo
        Case 1
            Call Grupo1
        Case 2
            Call Grupo2
        Case 3
            Call Grupo3
    End Select
    
    For i = 1 To Len(Letra)
        npeg = Mid(Letra, i, 1)
        If Not IsNumeric(npeg) Then
                For ix = 0 To 99
                    If N(Grupo, ix) = npeg Then
                        texto = texto & IIf(Len(Trim(ix)) = 1, "0" & ix, ix)
                        Exit For
                    End If
                Next
            Else
                ''Debug.Print "Letra: " & Letra & " //  npeg: " & npeg
                ''Debug.Print texto & " - Ant: " & Mid(Letra, IIf(i = 1, i, i - 1), 1) & " // " & "Suc: " & Mid(Letra, i + 1, 1)
                If IsNumeric(Mid(Letra, IIf(i = 1, i, i - 1), 1)) Then
                        texto = texto & npeg
                    Else
                        If IsNumeric(Mid(Letra, i + 1, 1)) Or Mid(Letra, i + 1, 1) = "" Then
                                texto = texto & npeg
                            Else
                                texto = texto & IIf(Len(npeg) = 1, "0" & npeg, npeg)
                        End If
                End If
        End If
    Next
    ConvLetra = texto
End Function
Private Function ConvNum(Grupo As String, Num As String)
    Dim npeg As String
    Dim texto As String
    Select Case Grupo
        Case 1
            Call Grupo1
        Case 2
            Call Grupo2
        Case 3
            Call Grupo3
    End Select
    
    For i = 1 To Len(Num) Step 2
        npeg = Mid(Num, i, 2) 'CInt(Mid(Num, i, 2))
        If Not N(Grupo, CInt(npeg)) = "" Then
                If Len(npeg) = 1 Then
                        texto = texto & npeg
                    Else
                        texto = texto & N(Grupo, npeg)
                End If
            Else
                If Mid(Num, i + 1, 2) = "" Then
                        texto = texto & npeg
                    Else
                        texto = texto & IIf(Len(npeg) = 1, "0" & npeg, npeg)
                End If
        End If
    Next
    ConvNum = texto
End Function

Private Sub Grupo1()
    'Grupo I
    N(1, 0) = "B"
    N(1, 2) = "W"
    N(1, 4) = "s"
    N(1, 6) = "u"
    N(1, 8) = "F"
    N(1, 10) = "H"
    N(1, 12) = "r"
    N(1, 14) = "I"
    N(1, 16) = "L"
    N(1, 18) = "t"
    N(1, 20) = "M"
    N(1, 22) = "O"
    N(1, 24) = "A"
    N(1, 26) = "P"
    N(1, 28) = "b"
    N(1, 30) = "Q"
    N(1, 32) = "R"
    N(1, 34) = "S"
    N(1, 36) = "U"
    N(1, 38) = "z"
    N(1, 40) = "V"
    N(1, 42) = "X"
    N(1, 44) = "Y"
    N(1, 46) = "Z"
    N(1, 48) = "a"
    N(1, 50) = "G"
    N(1, 52) = "c"
    N(1, 54) = "d"
    N(1, 56) = "e"
    N(1, 58) = "f"
    N(1, 60) = "y"
    N(1, 62) = "g"
    N(1, 64) = "h"
    N(1, 66) = "i"
    N(1, 68) = "j"
    N(1, 70) = "k"
    N(1, 72) = "D"
    N(1, 74) = "l"
    N(1, 76) = "E"
    N(1, 78) = "m"
    N(1, 80) = "n"
    N(1, 82) = "C"
    N(1, 84) = "o"
    N(1, 86) = "p"
    N(1, 88) = "q"
    N(1, 90) = "J"
    N(1, 92) = "K"
    N(1, 94) = "v"
    N(1, 96) = "w"
    N(1, 98) = "T"
    N(1, 99) = "x"
    N(1, 3) = "N"
End Sub
Private Sub Grupo2()
    'Grupo II
    N(2, 0) = "Y"
    N(2, 1) = "a"
    N(2, 3) = "G"
    N(2, 5) = "c"
    N(2, 7) = "d"
    N(2, 9) = "e"
    N(2, 11) = "f"
    N(2, 13) = "y"
    N(2, 15) = "g"
    N(2, 17) = "h"
    N(2, 19) = "i"
    N(2, 21) = "j"
    N(2, 23) = "k"
    N(2, 25) = "D"
    N(2, 27) = "l"
    N(2, 29) = "E"
    N(2, 31) = "m"
    N(2, 33) = "n"
    N(2, 35) = "C"
    N(2, 37) = "o"
    N(2, 39) = "p"
    N(2, 41) = "q"
    N(2, 43) = "J"
    N(2, 45) = "K"
    N(2, 47) = "v"
    N(2, 49) = "w"
    N(2, 51) = "T"
    N(2, 53) = "x"
    N(2, 55) = "N"
    N(2, 57) = "B"
    N(2, 59) = "W"
    N(2, 61) = "s"
    N(2, 63) = "u"
    N(2, 65) = "F"
    N(2, 67) = "H"
    N(2, 69) = "r"
    N(2, 71) = "I"
    N(2, 73) = "L"
    N(2, 75) = "t"
    N(2, 77) = "M"
    N(2, 79) = "O"
    N(2, 81) = "A"
    N(2, 83) = "P"
    N(2, 85) = "b"
    N(2, 87) = "Q"
    N(2, 89) = "R"
    N(2, 91) = "S"
    N(2, 93) = "U"
    N(2, 95) = "z"
    N(2, 97) = "V"
    N(2, 99) = "X"
    N(2, 90) = "Z"
End Sub
Private Sub Grupo3()
    'Grupo III
    N(3, 0) = "B"
    N(3, 3) = "W"
    N(3, 6) = "s"
    N(3, 9) = "u"
    N(3, 12) = "F"
    N(3, 15) = "H"
    N(3, 18) = "r"
    N(3, 21) = "I"
    N(3, 24) = "L"
    N(3, 27) = "t"
    N(3, 30) = "M"
    N(3, 33) = "O"
    N(3, 36) = "A"
    N(3, 39) = "P"
    N(3, 42) = "b"
    N(3, 45) = "Q"
    N(3, 48) = "R"
    N(3, 51) = "S"
    N(3, 54) = "U"
    N(3, 57) = "E"
    N(3, 59) = "m"
    N(3, 62) = "n"
    N(3, 65) = "C"
    N(3, 68) = "o"
    N(3, 71) = "p"
    N(3, 74) = "q"
    N(3, 77) = "J"
    N(3, 80) = "K"
    N(3, 83) = "v"
    N(3, 86) = "w"
    N(3, 89) = "T"
    N(3, 91) = "x"
    N(3, 94) = "N"
    N(3, 97) = "z"
    N(3, 2) = "V"
    N(3, 4) = "X"
    N(3, 6) = "Y"
    N(3, 8) = "Z"
    N(3, 10) = "a"
    N(3, 13) = "G"
    N(3, 16) = "c"
    N(3, 96) = "d"
    N(3, 76) = "e"
    N(3, 63) = "f"
    N(3, 82) = "y"
    N(3, 29) = "g"
    N(3, 55) = "h"
    N(3, 81) = "i"
    N(3, 73) = "j"
    N(3, 37) = "k"
    N(3, 52) = "D"
    N(3, 99) = "l"
End Sub

Private Function NGrupo(Num As String)
    Select Case Num
        Case 0
            NGrupo = 3
        Case 1
            NGrupo = 2
        Case 2
            NGrupo = 1
        Case 3
            NGrupo = 3
        Case 4
            NGrupo = 2
        Case 5
            NGrupo = 1
        Case 6
            NGrupo = 3
        Case 7
            NGrupo = 2
        Case 8
            NGrupo = 1
        Case 9
            NGrupo = 3
    End Select
    
    'If Num >= 0 And Num <= 3 Then
    '        NGrupo = 1
    '    Else
    '        If Num >= 4 And Num <= 7 Then
    '                NGrupo = 2
    '            Else
    '                If Num >= 8 And Num <= 9 Then
    '                    NGrupo = 3
    '                End If
    '        End If
    'End If
End Function
