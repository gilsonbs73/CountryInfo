VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CountryVO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private sISOCode As String
Private sName As String
Private sCapitalCity As String
Private sPhoneCode As String
Private sContinentCode As String
Private sCurrencyISOCode As String
Private sCountryFlag As String
Private cLanguages As New Collection

Public Property Get ISOCode() As String
    ISOCode = sISOCode
End Property

Public Property Get Name() As String
    Name = sName
End Property

Public Property Get CapitalCity() As String
    CapitalCity = sCapitalCity
End Property

Public Property Get PhoneCode() As String
    PhoneCode = sPhoneCode
End Property

Public Property Get ContinentCode() As String
    ContinentCode = sContinentCode
End Property

Public Property Get CurrencyISOCode() As String
    CurrencyISOCode = sCurrencyISOCode
End Property


Public Property Get CountryFlag() As String
    CountryFlag = sCountryFlag
End Property

Public Property Get Languages() As Collection
    Set Languages = cLanguages
End Property

Public Property Let ISOCode(pISOCode As String)
    sISOCode = pISOCode
End Property

Public Property Let Name(pName As String)
    sName = pName
End Property

Public Property Let CapitalCity(pCapitalCity As String)
    sCapitalCity = pCapitalCity
End Property

Public Property Let PhoneCode(pPhoneCode As String)
    sPhoneCode = pPhoneCode
End Property

Public Property Let ContinentCode(pContinentCode As String)
    sContinentCode = pContinentCode
End Property

Public Property Let CurrencyISOCode(pCurrencyISOCode As String)
    sCurrencyISOCode = pCurrencyISOCode
End Property

Public Property Let CountryFlag(pCountryFlag As String)
    sCountryFlag = pCountryFlag
End Property

Public Property Get CountryText() As String

Dim serializeObj As String


serializeObj = "ISOCode: " + sISOCode + vbNewLine
serializeObj = serializeObj + vbTab + "CountryName: " + sName + vbNewLine
serializeObj = serializeObj + vbTab + "CapitalCity: " + sCapitalCity + vbNewLine
serializeObj = serializeObj + vbTab + "PhoneCode: " + sPhoneCode + vbNewLine
serializeObj = serializeObj + vbTab + "ContinentCode: " + sContinentCode + vbNewLine
serializeObj = serializeObj + vbTab + "CurrencyISOCode: " + sCurrencyISOCode + vbNewLine
serializeObj = serializeObj + vbTab + "CountryFlag: " + sCountryFlag + vbNewLine
serializeObj = serializeObj + vbTab + "Langungues: " + vbNewLine

For Each languagues In cLanguages
    serializeObj = serializeObj + vbTab + vbTab + languagues.LanguagueText
Next languagues

CountryText = serializeObj
End Property
