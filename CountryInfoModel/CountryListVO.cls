VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CountryListVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private sSourceXML As String
Private cCountryList As New Collection


Public Property Get SourceXML() As String
    SourceXML = sSourceXML
End Property

Public Property Get CountryList() As Collection
    Set CountryList = cCountryList
End Property

Public Property Let SourceXML(pSourceXML As String)
    sSourceXML = pSourceXML
End Property


Public Property Get CountryText() As String

Dim serializeObj As String


serializeObj = "Resultado da Consulta" + vbNewLine
For Each countries In cCountryList
    serializeObj = serializeObj + "===================================" + vbNewLine
    serializeObj = serializeObj + countries.CountryText
Next countries


CountryText = serializeObj
End Property
