Attribute VB_Name = "mdlFuncoes"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long

Public Function ReadIniFile(ByVal strINIFile As String, ByVal strSECTION As String, ByVal strKey As String) As String
'Função para captura e leitura do arquivo INI
    Dim strBuffer As String
    Dim intPos As Integer
    strBuffer = Space$(gintMAX_SIZE)
    If GetPrivateProfileString(strSECTION, strKey, "", strBuffer, gintMAX_SIZE, strINIFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = ""
    End If
End Function


Public Function StripTerminator(ByVal strString As String) As String
'Função para captura e leitura do arquivo INI
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
