VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CountryLanguagesVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
 Private sISOCode As String
 Private sName As String
 
Public Property Get Name() As String
    Name = sName
End Property

Public Property Get ISOCode() As String
    ISOCode = sISOCode
End Property

Public Property Let Name(pName As String)
    sName = pName
End Property

Public Property Let ISOCode(pISOCode As String)
    sISOCode = pISOCode
End Property

Public Property Get LanguagueText() As String

Dim serializeObj As String

serializeObj = "---------------------------" + vbNewLine
serializeObj = serializeObj + vbTab + vbTab + "ISOCode: " + sISOCode + vbNewLine
serializeObj = serializeObj + vbTab + vbTab + "LanguageName: " + sName + vbNewLine

LanguagueText = serializeObj

End Property

