Attribute VB_Name = "Genio"
Option Explicit
Private Genio As IAgentCtlCharacter
Const DadosGenio = "Genie.acs"
Public Sub LoadGenio()
On Error GoTo TratErro:
    MDIForm_Main.Agent.Characters.Load "Genio", App.Path & "\" & DadosGenio
    Set Genio = MDIForm_Main.Agent.Characters("Genio")
    Exit Sub
TratErro:
    Call RegLogErros(Err.Number, Err.Description, "Modulo_Genio/LoadGenio", UsuarioID)
    MsgBox Err.Description, vbInformation, Err.Number
End Sub
Public Sub CAgent(Texto As String)
    On Error GoTo TratErro1
    Dim fala As String
    'Configuracoes
    Genio.SoundEffectsOn = True
    '***************************
    Genio.MoveTo 400, 300
    Genio.Show
    Genio.Play "Greet"
    'Genio.MoveTo 400, 100
    fala = "\pit=1\\spd=1\" & Texto & "\pau=1\"
    Genio.Speak fala
    'Genio.MoveTo 690, 450
    'Genio.Play ("domagic1")
    Genio.Play "wave"
    'Genio.Play ("domagic2")
    'Genio.Play "Hide"
    Genio.Hide
    'Genio.Stop
    'Modificadores do discurso :
'   1.  \emp\       enfatiza a palavra
'   2.  \pau = m\   pause de m milisegundos
'   3.  \pit = p\    voz para p Hertz (1 - 400)
'   4.  \spd = s\  define a velocidade para s palavras por minuto
    Exit Sub
TratErro1:
    Call RegLogErros(Err.Number, Err.Description, "Modulo_Genio/CAgent", UsuarioID)
    MsgBox Texto, vbInformation, "PGE - Aviso"
End Sub
