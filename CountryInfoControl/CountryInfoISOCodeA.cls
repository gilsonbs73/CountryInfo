VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CountryInfoISOCode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Private filteredCountries As New CountryListVO


Public Function GetCountryIsoCode(URI As String, FilteredIsoCode As String)
    Dim resultSet As New DOMDocument60
    Dim countryModel As New CountryInfoModel.FullCountryInfo
    Dim resultApiCountryList As CountryListVO
    Dim countryResult As Variant
        
    
    Dim APITextResponse As String

    If URI = "" Then URI = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries"
    
    Set filteredCountries = New CountryListVO
    
    Set resultApiCountryList = countryModel.GetFullCountryInfo(URI)
    
    
    'Filtrar somente ISOCode A
     For Each countryResult In resultApiCountryList.countryList
        If UCase(Left(countryResult.ISOCode, 1)) = UCase(FilteredIsoCode) Then
            filteredCountries.countryList.Add countryResult
        End If
     Next countryResult

     GetCountryIsoCode = filteredCountries.CountryText
     
End Function

Public Function GetCountryIsoCodeXML(URI As String, FilteredIsoCode As String)
    Dim resultSet As New DOMDocument60
    Dim countryModel As New CountryInfoModel.FullCountryInfo
    Dim APITextResponse As String

    If URI = "" Then URI = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries"
          
    'APITextResponse = countryModel.GetFullCountryInfo(URI)
    APITextResponse = countryModel.GetFullCountryInfoXML(URI)
    
    resultSet.loadXML (APITextResponse)

    GetCountryIsoCodeXML = APITextResponse
    DoEvents

End Function

Public Sub SaveFilteredData()
    Dim countryData As New FullCountryInfo
    
    If Not IsNull(filteredCountries) Then
        Call countryData.SaveData(filteredCountries)
    End If
End Sub




