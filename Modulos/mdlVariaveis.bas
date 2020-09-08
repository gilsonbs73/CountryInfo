Attribute VB_Name = "mdlVariaveis"
Public serverMySQL As String
Public portaMySQL As String
Public bdMySQL As String
Public userMySQL As String
Public senhaMySQL As String

Public con As New ADODB.Connection
Public rs As ADODB.Recordset

Public Const gintMAX_SIZE = 255

Public strSql As String
