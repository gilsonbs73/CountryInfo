Attribute VB_Name = "mdlIniciar"
Option Explicit
Sub Main()
    Dim con_str As String
    
    On Error GoTo TratarErro
    serverMySQL = Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "SERVIDOR"))
    portaMySQL = Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "PORTA"))
    bdMySQL = Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "BANCODADOS"))
    userMySQL = Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "USER"))
    senhaMySQL = Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "SENHA"))
    'senhaMySQL = ""
    '
    con_str = "DRIVER={MySQL ODBC 3.51 Driver};" _
        & "SERVER=" & serverMySQL & ";" _
        & "PORT=" & portaMySQL & ";" _
        & "DATABASE=" & bdMySQL & ";" _
        & "UID=" & userMySQL & ";PWD=" & senhaMySQL & "; OPTION= 1 + 2 + 8 + 32 + 2048 + 16384"
        
        With con
            .CursorLocation = adUseClient
            .ConnectionString = con_str
            .Open con_str
        End With
TratarErro:
    If Err.Number <> 0 Then
       MsgBox ("Erro na conexão com o servidor ")
       End
    Else
        CountryInfoWithISOCodeA.Show
    End If
End Sub


