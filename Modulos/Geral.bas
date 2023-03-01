Attribute VB_Name = "Geral"
Option Explicit

Public VersaoAno    As String 'Ano da Versao do Sistema


Public Usuario      As String    'Chave do Usuario
Public UsuarioID    As String  'Usuario Identificador


Public BD           As Database
Public BD_ADO       As New ADODB.Connection

Public PathServ     As String
Public PathBD       As String
Public Chave        As String
'Public Acesso As String
'Public DMatr(23) As String


'CONFIGURAÇÃO DO SERVIDOR
'Dita se o programa rodando é servidor(1) ou cliente(0)
Public Status_Serv  As String
Dim IP_Serv         As String
Dim Port_Serv       As String
    

'Sistema de Grupos de Acesso ao Sistema
Public GrupoAcesso  As New Collection
Public GA(99, 5)    As Boolean

'
Dim RsUsuario       As Recordset
'Dim sBd             As Integer

Public TipoUsoSoftware As Boolean 'True - Execucao False - Demonstracao





Public Sub LoadLocalBD()
    Dim Arquivo     As String
     

    Arquivo = FreeFile
    Open App.path & "\CESNet.srv" For Input As Arquivo
    Line Input #Arquivo, Status_Serv
    Line Input #Arquivo, IP_Serv
    Line Input #Arquivo, Port_Serv
    
    
    Close #Arquivo
    'PathBD = Crypto(PathBD) & "\DATABASE\DADOS"
    Select Case Status_Serv
        Case 0
            PathServ = "\\" & IP_Serv
            PathBD = PathServ '& "\DATABASE\DADOS"
        Case 1
            PathServ = "\\" & IP_Serv
            PathBD = PathServ '& "\DATABASE\DADOS"
        Case 2
            PathServ = IP_Serv
            PathBD = PathServ
    End Select
    
End Sub

Public Sub Main()
    'On Error GoTo TrtErro
    'If App.PrevInstance = True Then
    '    MsgBox "O CESNet já esta sendo executado...", vbInformation, "CESNet - Atenção!"
    '    End
    'End If
   ' Dim command As String
    If Command <> "" Then
        Select Case UCase(Command)
            Case "UDB" 'Update Database
                Form_SQLExecuteAuto.CarregarFormulario
            Case "MANUTENCAO"
                Form_SQLDatabase.Show
                Exit Sub
            Case "REPARAR"
                RepararBD
            Case "COMPACTAR"
                CompactarBD
            Case "?"
                MsgBox "Use:" & vbCrLf & _
                       ">CESNET UDB (Atualiza as tabelas da base de dados)" & vbCrLf & _
                       ">CESNET COMPACTAR (Compacta a base de dados)" & vbCrLf & _
                       ">CESNET REPARAR (Repara a base de dados)", _
                       vbInformation, App.EXEName
                Exit Sub
            Case Else
                MsgBox Command & " - comando invalido!"
                 End
        End Select
    End If
    '**********************************************************************************
    '**********************************************************************************
    '*** Objetivo: Limita o uso de até 10 registros por tabela
    '*** Variavel: TipoUsoSoftware
    '*** Opcoes  : True - Execucao
    '***           False - Demonstracao
    TipoUsoSoftware = True
    '**********************************************************************************
    '**********************************************************************************
    
    VersaoAno = "2013"
    PgSoftwareID
    Call AbrirBD_DAO
    Call PgRegrasSis
    
    Form_Splash.Show 1
    
    If Form_Login.LoadLogin = False Then
        BD.Close
        End
    End If
    
    MDIForm_Main.Show
    Exit Sub
TrtErro:
    Call RegLogErros(Err.Number, Err.Description, "MAIN", UsuarioID)
    MsgBox "- Descrição: " & Err.Description, vbCritical, "Erro n.: " & Err.Number
    End
End Sub
Public Function CarregarServidor()
    MDIForm_Main.Winsock_Main.LocalPort = Port_Serv
    MDIForm_Main.Winsock_Main.Listen
End Function
Public Sub AbrirBD_DAO()
 On Error GoTo TratErro
'    Dim Arquivo     As String
''

'    Arquivo = FreeFile
'    Open App.path & "\CESNet.srv" For Input As Arquivo
'    Line Input #Arquivo, Status_Serv
'    Line Input #Arquivo, IP_Serv
'    Line Input #Arquivo, Port_Serv
'
'
'    Close #Arquivo
'    'PathBD = Crypto(PathBD) & "\DATABASE\DADOS"
'    Select Case Status_Serv
'        Case 0
'            PathServ = "\\" & IP_Serv
'            PathBD = PathServ '& "\DATABASE\DADOS"
'        Case 1
'            PathServ = "\\" & IP_Serv
'            PathBD = PathServ '& "\DATABASE\DADOS"
'        Case 2
'            PathServ = IP_Serv
'            PathBD = PathServ
'    End Select
    LoadLocalBD
    Set BD = DBEngine.Workspaces(0).OpenDatabase(PathBD & "\Database\Dados\Dados.mdb", False, False, ";PWD=k3bw82")

    Exit Sub
TratErro:
    Call RegLogErros(Err.Number, Err.Description, "Modulo_Geral: Erro ao abrir Banco de dados.", UsuarioID)
    Select Case Err.Number
        Case 3024, 3044, 53 'Banco de Dados não encontrado
            MsgBox "- Erro no acesso ao Banco de Dados." & Chr(13) & "  Caso o problema persista chame o técnico.", vbCritical, "Erro de Acesso - n.º " & Err.Number
            Form_ConectarBD.Show 1
            Form_Login.Hide
        'Case 3224 'O erro 3224 significa que o Banco de Dados (BD) foi corrompido
        '    MsgBox "- Erro no acesso ao Banco de Dados, o sistema tentará repara-lo." & Chr(13) & " Caso o problema persista chame o técnico.", vbCritical, "Erro de Acesso - n.º " & Err.Number
        '    DBEngine.RepairDatabase (PathBD & "\Dados.mdb") 'Repara o BD
        '    DBEngine.CompactDatabase PathBD & "\Dados.mdb", PathBD & "\DadosBKP.mdb" 'Compacta o Banco de Dados Reparado e Renomeia.
        '    'É importante compactar o BD apos repara-lo devido ao aumento do tamanho do mdb
        '    Kill (PathBD & "\Dados.mdb") 'Apaga o BD antigo
        '    Name PathBD & "\DadosBKP.mdb" As PathBD & "\Dados.mdb" 'Renomeia o BD reparado para o corrente no sistema
        '    MsgBox "Correção concluida. O sistema será encerrado", vbExclamation, "CESNet - Aviso"
        '    End
        Case 3043
            MsgBox "Erro ao localizar o SERVIDOR. Favor checar sua rede ou se o servidor mudou de IP.", vbCritical, "CESNet - Aviso"
            End
        Case Else
            MsgBox "- Modulo: AbrirBD_DAO" & Chr(13) & "- Descrição: " & Err.Description & Chr(13) & "- O sistema será encerrado." & Chr(13) & "- Chame o Técnico.", vbCritical, "Erro de Acesso - n.º " & Err.Number
            End
    End Select
End Sub

Public Sub AbrirBD_ADO()
 On Error GoTo TratErro
   LoadLocalBD
    
    BD_ADO.Open "Provider=Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source=" & PathBD & "\Database\Dados\Dados.mdb;" & _
            "Jet OLEDB:Database Password=k3bw82;"

    Exit Sub
TratErro:
    Call RegLogErros(Err.Number, Err.Description, "Modulo_Geral: Erro ao abrir Banco de dados.", UsuarioID)
    Select Case Err.Number
        Case 3024, 3044, 53 'Banco de Dados não encontrado
            MsgBox "- Erro no acesso ao Banco de Dados." & Chr(13) & "  Caso o problema persista chame o técnico.", vbCritical, "Erro de Acesso - n.º " & Err.Number
            Form_ConectarBD.Show 1
            Form_Login.Hide
        Case 3043
            MsgBox "Erro ao localizar o SERVIDOR. Favor checar sua rede ou se o servidor mudou de IP.", vbCritical, "CESNet - Aviso"
            End
        Case Else
            MsgBox "- Modulo: AbrirBD_ADO" & Chr(13) & "- Descrição: " & Err.Description & Chr(13) & "- O sistema será encerrado." & Chr(13) & "- Chame o Técnico.", vbCritical, "Erro de Acesso - n.º " & Err.Number
            End
    End Select
End Sub

Function Crypto(Texto As String)
    Dim Chave As String, PosS As Long, PosC As Long, TempString As String
    Chave = "1028KLP"
    'Chave = "83K12L7P0"
    TempString = space$(Len(Texto))
    PosC = 1
    For PosS = 1 To Len(Texto)
        If PosC > Len(Chave) Then PosC = 1
        Mid(TempString, PosS, 1) = Chr$(Asc(Mid(Texto, PosS, 1)) Xor Asc(Mid(Chave, PosC, 1)))
        If Asc(Mid(TempString, PosS, 1)) = 0 Then Mid(TempString, PosS, 1) = Mid(Texto, PosS, 1)
        PosC = PosC + 1
    Next PosS
    Crypto = TempString
End Function
Public Sub RegLogErros(ByVal Num As String, ByVal Descr As String, ByVal Form As String, ByVal Usuario As String)
'Colocar o txt abaixo em todos On Error
'RegLogErros Err.Number, Err.Description, "MODULO", 0

'define o ObjPreview filesystem e demais variaveis
Dim fso As New FileSystemObject
Dim Arquivo As File
Dim arquivoLog As TextStream
Dim msg As String
Dim caminho As String
    caminho = App.path & "\ErrLog.txt"
'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
'monta informações para gerar a linha com erro
    msg = "[" & Now & "]" & Form & ":[" & Num & "-" & Descr & "]"
' inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)
'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
End Sub
Public Sub RegLog(ByVal Matr As String, ByVal descricao As String)
On Error GoTo TrtErro
'Registra oq os usuarios do sistema estao fazendo

'define o ObjPreview filesystem e demais variaveis
Dim fso As New FileSystemObject
Dim Arquivo As File
Dim arquivoLog As TextStream
Dim msg As String
Dim caminho As String
    caminho = PathServ & "\Database\Log\" & Format(Date, "yyyymm") & ".txt"
'se o arquivo não existir então cria
    If fso.FileExists(caminho) Then
            Set Arquivo = fso.GetFile(caminho)
        Else
            Set arquivoLog = fso.CreateTextFile(caminho)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(caminho)
    End If
'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
'monta informações para gerar a linha da mensagem
    msg = "[" & Now & "] - " & Usuario & left(String(10, " "), 10 - Len(Trim(Usuario))) & "[" & IIf(IsNull(Matr), "00.000.0000", Matr) & "] = " & UCase(descricao)
' inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)
'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing
    Exit Sub
TrtErro:
    Call RegLogErros(Err.Number, Err.Description, "Modulo_Geral - Erro ao gravar Log.", UsuarioID)
End Sub


Public Sub RepararBD()
    
    'Dim RsTMP As Recordset
    On Error GoTo TratErro:
    'BD.Close
    LoadLocalBD
    'DBEngine.RepairDatabase "Northwind.mdb"
    DBEngine.RepairDatabase PathBD & "\Database\Dados\Dados.mdb" ', PathBD & "\Database\Dados\DadosBKP.mdb", , , ";PWD=k3bw82" 'Compacta o Banco de Dados Reparado e Renomeia.
    'É importante compactar o BD apos repara-lo devido ao aumento do tamanho do mdb
    'Kill (PathBD & "\Database\Dados\Dados.mdb") 'Apaga o BD antigo
    'Name PathBD & "\Database\Dados\DadosBKP.mdb" As PathBD & "\Database\Dados\Dados.mdb" 'Renomeia o BD reparado para o corrente no sistema
    
    'AbrirBD_DAO
    'Set RsTMP = BD.OpenRecordset("SELECT * FROM Config")
    'sTMP.MoveFirst
    'RsTMP.Edit
    'RsTMP.Fields("UltManutencao") = Date
    'RsTMP.Update
    'RsTMP.Close
    
    
    MsgBox "Correção concluida. Banco de Dados Restaurado.", vbExclamation, "AVISO"
    
    'BD.Close
    
    Call RegLog(UsuarioID, "REPARACAO DE SEGURANCA DO BANCO DE DADOS")
    Exit Sub
TratErro:
    RegLogErros Err.Number, Err.Description, "RepararBD", 0
    'Call RegLogerros UsuarioID, "REPARACAO DE SEGURANCA DO BANCO DE DADOS")
    MsgBox Err.Description & vbCrLf & "Operação cancelada.", vbInformation, "Erro n. " & Err.Number
End Sub

Public Sub StatusBD()

    On Error GoTo tErroSBD
    Dim m As String
    m = FileLen(PathBD & "\Database\Dados\Dados.mdb")
    MDIForm_Main.StatusBar_Menu.Panels(8).Picture = MDIForm_Main.IL_Main16x16.ListImages.Item(7).Picture
    'Codigo abaixo retirado devido dar erro na iniciação do sistema no Timer_StatusBD
    'Call PgRegrasSis
    Exit Sub
tErroSBD:
    MDIForm_Main.StatusBar_Menu.Panels(8).Picture = MDIForm_Main.IL_Main16x16.ListImages.Item(8).Picture
    Exit Sub
End Sub
Public Sub LoadStatusBarr(Optional sBd = 8)
    On Error GoTo TratErroSB
    With MDIForm_Main.StatusBar_Menu
        .Panels(1) = UnidadeEnsino & " - " & UnidadeEnsinoNome
        .Panels(2) = "Usuario: " & Usuario
        .Panels(7) = "IP: " & MDIForm_Main.Winsock_Main.LocalIP
        .Panels(8).Width = Val(MDIForm_Main.IL_Main16x16.ListImages.Item(sBd).Picture.Width)
        .Panels(8).Picture = MDIForm_Main.IL_Main16x16.ListImages.Item(sBd).Picture
        .Panels(9).Alignment = sbrCenter
        '.Panels(9).Width = FontWidth("CESNet - v." & Versao)
        .Panels(9).Text = "CESNet - v." & Versao
        .Panels(1).Width = Val(Screen.Width) - (Val(Val(.Panels(2).Width) + Val(.Panels(3).Width) + Val(.Panels(4).Width) + Val(.Panels(5).Width) + Val(.Panels(6).Width) + Val(.Panels(7).Width) + Val(.Panels(8).Width) + Val(.Panels(9).Width) + 200))
    End With
    Exit Sub
TratErroSB:
    MsgBox "Erro ao carregar barra de status." & Chr(13) & Err.Description, vbInformation, "Erro n. " & Err.Number
    Resume Next
End Sub

Public Sub LoadGrupoUsu(GrupoID As Integer)

    
    Dim RsGrupoForm As Recordset
    Dim cont As Integer
    Set GrupoAcesso = Nothing
    cont = 0
    

    
    Set RsGrupoForm = BD.OpenRecordset("SELECT * FROM UsuGrupoForm WHERE GrupoID = " & GrupoID)
    If RsGrupoForm.BOF And RsGrupoForm.EOF Then
            MsgBox "Erro ao localizar o Grupo de Usuarios. Operação cancelada.", vbInformation, "CESNet - Aviso"
            Exit Sub
        Else
            RsGrupoForm.MoveFirst
            Do Until RsGrupoForm.EOF
                cont = cont + 1
                GrupoAcesso.Add cont, RsGrupoForm.Fields("Form")
                GA(cont, 1) = RsGrupoForm.Fields("Inc")
                GA(cont, 2) = RsGrupoForm.Fields("Alt")
                GA(cont, 3) = RsGrupoForm.Fields("Exc")
                GA(cont, 4) = RsGrupoForm.Fields("Impr")
                GA(cont, 5) = RsGrupoForm.Fields("Cons")
                RsGrupoForm.MoveNext
            Loop
            RsGrupoForm.Close
    End If
End Sub

Public Function ChkAcesso(Formulario As String, Acao As String) As Boolean
    On Error GoTo TratAcesso
    Select Case UCase(Acao)
        Case "N" 'Novo
            Acao = 1
        Case "A" 'Alterar
            Acao = 2
        Case "E" 'xcluir
            Acao = 3
        Case "I" 'mprimir
            Acao = 4
        Case "C" 'onsultar
            Acao = 5
    End Select
    If GA(GrupoAcesso(Formulario), Acao) = True Then
            ChkAcesso = True
        Else
            MsgBox "Usuário(a) " & Usuario & " seu acesso a esta operação foi negado(a).", vbExclamation, "CESNet - Sistema Segurança"
            ChkAcesso = False
    End If
    Exit Function
TratAcesso:
    If Err.Number = 5 Then
        Call RegLogErros(Err.Number, "Acesso - Erro ao localizar ID do formulario", "ChkAcesso", UsuarioID)
        MsgBox "Erro ao localizar o ID do formulário solicitado." & Chr(13) & _
                "Formulário bloqueado por questão de segurança." & Chr(13) & _
                "Por favor avise ao suporte!", vbCritical, "CESNet - Aviso!"
        Else
            Call RegLogErros(Err.Number, Err.Description, "ChkAcesso", UsuarioID)
            MsgBox Err.Description & Chr(13) & _
                "Formulário bloqueado por questão de segurança." & Chr(13) & _
                "Por favor avise ao suporte!", vbCritical, "CESNet - Erro n." & Err.Number
    End If
    ChkAcesso = False
    'Resume Next
End Function
'Public Sub AtualizarVersao()
'    Shell PathBD & "\Database\PRG\SETUP.EXE"
'End Sub
Public Sub CompactarBD()
    Dim RsTMP As Recordset
    On Error GoTo TratErro:
    'BD.Close
    LoadLocalBD
    DBEngine.CompactDatabase PathBD & "\Database\Dados\Dados.mdb", PathBD & "\Database\Dados\DadosBKP.mdb", , , ";PWD=k3bw82" 'Compacta o Banco de Dados Reparado e Renomeia.
    'É importante compactar o BD apos repara-lo devido ao aumento do tamanho do mdb
    Kill (PathBD & "\Database\Dados\Dados.mdb") 'Apaga o BD antigo
    Name PathBD & "\Database\Dados\DadosBKP.mdb" As PathBD & "\Database\Dados\Dados.mdb" 'Renomeia o BD reparado para o corrente no sistema
    
    AbrirBD_DAO
    Set RsTMP = BD.OpenRecordset("SELECT * FROM Config")
    RsTMP.MoveFirst
    RsTMP.Edit
    RsTMP.Fields("UltManutencao") = Date
    RsTMP.Update
    RsTMP.Close
    
    
    MsgBox "Correção concluida. Banco de Dados Restaurado.", vbExclamation, "AVISO"
    
    BD.Close
    
    Call RegLog(UsuarioID, "COMPACTACAO DE SEGURANCA DO BANCO DE DADOS")
     'Call RegLog(UsuarioID, "COMPACTACAO DE SEGURANCA DO BANCO DE DADOS")
    Exit Sub
TratErro:
    'If Err.Number = 91 Then Resume Next
    RegLogErros Err.Number, Err.Description, "CompactarBD", 0
    MsgBox Err.Description & vbCrLf & "Operação cancelada.", vbInformation, "Erro n. " & Err.Number
End Sub


Public Sub ExpArq(locArq As String, msg As String)
'define o ObjPreview filesystem e demais variaveis
Dim fso As New FileSystemObject
Dim Arquivo As File
Dim arquivoLog As TextStream
'Dim caminho2  As String
    'caminho2 = caminho & ".xls"

    'caminho = caminho & "\40253676.txt"
'se o arquivo não existir então cria
    If fso.FileExists(locArq) Then
            Set Arquivo = fso.GetFile(locArq)
        Else
            Set arquivoLog = fso.CreateTextFile(locArq)
            arquivoLog.Close
            Set Arquivo = fso.GetFile(locArq)
    End If
'prepara o arquivo para anexa os dados
    Set arquivoLog = Arquivo.OpenAsTextStream(ForAppending)
'monta informações para gerar a linha com erro
    'msg = "[" & Now & "]" & Form & ":[" & Num & "-" & Descr & "]"
' inclui linhas no arquivo texto
    arquivoLog.WriteLine msg
   
    ''Debug.Print msg & " - " & Len(msg)
' escreve uma linha em branco no arquivo - se voce quiser
'arquivoLog.WriteBlankLines (1)
'fecha e libera o ObjPreview
    arquivoLog.Close
    Set arquivoLog = Nothing
    Set fso = Nothing

End Sub

Public Function rc(sTexto As String) As String
    rc = Replace(sTexto, "'", "''")
End Function
Public Function cNull(sTexto As Variant) As String
    If IsNull(sTexto) Then
            cNull = ""
        Else
            cNull = sTexto
    End If
End Function
Public Function ValidarSoftware(Tabela As String) As Boolean

'**********************************************************************************
'**********************************************************************************
'*** Objetivo: Limita o uso de até 10 registros por tabela
'*** Variavel: TipoUsoSoftware
'*** Opcoes  : True - Execucao
'***           False - Demonstracao
'**********************************************************************************
'**********************************************************************************
    On Error Resume Next
    Dim Rst     As Recordset
    Dim sSQL    As String
    'TipoUsoSoftware As Boolean 'True - Execucao False - Demonstracao
    If TipoUsoSoftware = True Then
        ValidarSoftware = True
        Exit Function
    End If
    
    sSQL = "SELECT * FROM " & Tabela
    Set Rst = BD.OpenRecordset(sSQL)
    If Rst.BOF And Rst.EOF Then
            ValidarSoftware = True
        Else
            Rst.MoveLast
            If Rst.RecordCount > 10 Then
                    ValidarSoftware = False
                Else
                    ValidarSoftware = True
            End If
    End If
    Rst.Close
    If ValidarSoftware = False Then
        MsgBox "Versão de demonstração, valido somente para 10 (dez) registros.", vbInformation, "DEMONSTRAÇÃO"
    End If
End Function
