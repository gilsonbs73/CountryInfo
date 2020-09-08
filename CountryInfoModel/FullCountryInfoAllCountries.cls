VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FullCountryInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal LpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal lSize As Long, ByVal lpFilename As String) As Long

Public Function GetFullCountryInfoXML(URI As String)
    Dim Req2 As New WinHttp.WinHttpRequest
    Dim result As String
    
    'Conexão com a API de FullCountryInfoAllCountries para carregr todos os paises
        With Req2
            '.Open "GET", "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries", async:=False
            .Open "GET", URI, async:=False
            .Send
        End With
        'result = "<?xml version='1.0' encoding='UTF-8'?>"
        result = Req2.ResponseText
        'result = Replace(result, "utf-8", "ISO-8859-1")
        
        result = Req2.ResponseText
        GetFullCountryInfoXML = Req2.ResponseText
        
End Function

Public Function GetFullCountryInfo(URI As String)
    Dim Req2 As New WinHttp.WinHttpRequest
    Dim result As String
    
    'Conexão com a API de FullCountryInfoAllCountries para carregr todos os paises
        With Req2
            '.Open "GET", "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries", async:=False
            .Open "GET", URI, async:=False
            .Send
        End With
        'result = "<?xml version='1.0' encoding='UTF-8'?>"
        result = Req2.ResponseText
        'result = Replace(result, "utf-8", "ISO-8859-1")
        
        Set GetFullCountryInfo = LoadCountryObjectFromXML(result)  'Req2.ResponseText
        
End Function

Private Function LoadCountryObjectFromXML(pXML As String)

Dim doc As New MSXML2.DOMDocument
Dim nodeC As MSXML2.IXMLDOMNode
Dim nodeL As MSXML2.IXMLDOMNode
Dim success As Boolean
Dim nodeCountry As MSXML2.IXMLDOMNodeList
Dim nodeLanguage As MSXML2.IXMLDOMNodeList
Dim cCountryList As New CountryListVO
      
Dim country As CountryVO
Dim languanges As CountryLanguagesVO

success = doc.loadXML(pXML)
   
If success = False Then
  MsgBox doc.parseError.reason
Else
    
    Set nodeCountry = doc.selectNodes("/ArrayOftCountryInfo/tCountryInfo")
    cCountryList.SourceXML = pXML
    
    For Each nodeC In nodeCountry
        Set country = New CountryVO
        
        country.ISOCode = nodeC.selectSingleNode("sISOCode").Text
        country.Name = Replace(nodeC.selectSingleNode("sName").Text, "'", " ")
        country.CapitalCity = Replace(nodeC.selectSingleNode("sCapitalCity").Text, "'", " ")
        country.PhoneCode = nodeC.selectSingleNode("sPhoneCode").Text
        country.ContinentCode = nodeC.selectSingleNode("sContinentCode").Text
        country.CurrencyISOCode = nodeC.selectSingleNode("sCurrencyISOCode").Text
        country.CountryFlag = nodeC.selectSingleNode("sCountryFlag").Text
        
        For Each nodeL In nodeC.selectSingleNode("Languages").childNodes
            Set languanges = New CountryLanguagesVO
            
            languanges.ISOCode = nodeL.selectSingleNode("sISOCode").Text
            languanges.Name = Replace(nodeL.selectSingleNode("sName").Text, "'", " ")
            
            country.Languages.Add languanges
            'Call AddCollectionByVal(ByVal country.Languages, ByVal languanges)
        Next nodeL
        
        cCountryList.CountryList.Add country
        
        'Call AddCollectionByVal(ByVal cCountryList.CountryList, ByVal country)
        
     Next nodeC
     Set LoadCountryObjectFromXML = cCountryList
End If

End Function

Private Sub AddCollectionByVal(ByVal collectionRef As Collection, ByVal objectRef As Object)
    collectionRef.Add objectRef
End Sub

Public Sub SaveData(CountryList As CountryListVO)
    Dim countryInfo
    Dim listLanguange
    
    Dim sISOCode As String
    Dim sName As String
    Dim sCapitalCity
    Dim sPhoneCode
    Dim sContinentCode
    Dim sCurrencyISOCode
    Dim sCountryFlag
         
    Dim sISOCodeL As String
    Dim sNameL As String
         
    Dim idISOCode As String
    Dim idCurrency As Integer
    Dim idContinent As Integer
    Dim idISOCodeL As String
    '
    Dim serverMySQL As String
    Dim portaMySQL As String
    Dim bdMySQL As String
    Dim userMySQL As String
    Dim senhaMySQL As String
    
    Dim con As New ADODB.Connection
    Dim rs As ADODB.Recordset
    
    Dim gintMAX_SIZE
    Dim strSql As String
    gintMAX_SIZE = 255
    
    '
    Dim con_str As String
    
    On Error GoTo TratarErro
    serverMySQL = "localhost" ' Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "SERVIDOR"))
    portaMySQL = "3306" 'Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "PORTA"))
    bdMySQL = "dbcountryinfo" 'Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "BANCODADOS"))
    userMySQL = "root" 'Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "USER"))
    senhaMySQL = "" 'Trim(ReadIniFile(App.Path + "\Parametrosgerais.ini", "Server", "SENHA"))
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
    End If
    '
    For Each countryInfo In CountryList.CountryList
        sISOCode = countryInfo.ISOCode
        sName = countryInfo.Name
        sCapitalCity = countryInfo.CapitalCity
        sPhoneCode = countryInfo.PhoneCode
        sContinentCode = countryInfo.ContinentCode
        sCurrencyISOCode = countryInfo.CurrencyISOCode
        sCountryFlag = countryInfo.CountryFlag
        '
        'String para verificar se a moeda do pais existe, se existir carrega na variavel, se nao existir insere e pega o id
        strSql = "IF (SELECT 1 = 1 FROM tb_currency WHERE currency_iso_code='" & sCurrencyISOCode & "') THEN "
        strSql = strSql & " BEGIN "
        strSql = strSql & " SELECT id FROM tb_currency WHERE currency_iso_code='" & sCurrencyISOCode & "';"
        strSql = strSql & " END; "
        strSql = strSql & " ELSE "
        strSql = strSql & " BEGIN "
        strSql = strSql & " INSERT INTO tb_currency(currency_iso_code) VALUES('" & sCurrencyISOCode & "');"
        strSql = strSql & " SELECT LAST_INSERT_ID() AS id; "
        strSql = strSql & " END; "
        strSql = strSql & " END IF; "
        Set rs = con.Execute(strSql)
        If Not rs.EOF Then
            idCurrency = rs!id
        End If
        '
        'String para verificar se o continente do pais existe, se existir carrega na variavel, se nao existir insere e pega o id
        strSql = "IF (SELECT 1 = 1 FROM tb_continent WHERE continent_code='" & sContinentCode & "') THEN "
        strSql = strSql & " BEGIN "
        strSql = strSql & " SELECT id FROM tb_continent WHERE continent_code='" & sContinentCode & "';"
        strSql = strSql & " END; "
        strSql = strSql & " ELSE "
        strSql = strSql & " BEGIN "
        strSql = strSql & " INSERT INTO tb_continent(continent_code) VALUES('" & sContinentCode & "');"
        strSql = strSql & " SELECT LAST_INSERT_ID() AS id; "
        strSql = strSql & " END; "
        strSql = strSql & " END IF; "
        Set rs = con.Execute(strSql)
        If Not rs.EOF Then
            idContinent = rs!id
        End If
        '
        'String para gravar ou atualizar os dados do Pais.
        'sCapitalCity = Replace(nodeC.selectSingleNode("sCapitalCity").Text, "'", " ")
        strSql = "IF (SELECT 1 = 1 FROM tb_country WHERE sisocode='" & sISOCode & "') THEN "
        strSql = strSql & " BEGIN "
        strSql = strSql & " UPDATE tb_country SET sisocode='" & sISOCode & "'"
        strSql = strSql & " ,country_name='" & sName & "'"
        strSql = strSql & " ,capital_city='" & sCapitalCity & "'"
        strSql = strSql & " ,phone_code='" & sPhoneCode & "'"
        strSql = strSql & " ,country_flag='" & sCountryFlag & "'"
        strSql = strSql & " ,id_currency='" & idCurrency & "'"
        strSql = strSql & " ,id_continent='" & idContinent & "'"
        strSql = strSql & " WHERE sisocode='" & sISOCode & "';"
        strSql = strSql & " SELECT id FROM tb_country WHERE sisocode='" & sISOCode & "';"
        strSql = strSql & " END; "
        strSql = strSql & " ELSE "
        strSql = strSql & " BEGIN "
        strSql = strSql & " INSERT INTO tb_country(sisocode,country_name,capital_city,phone_code,country_flag"
        strSql = strSql & " ,id_currency,id_continent) "
        strSql = strSql & " VALUES('" & sISOCode & "','" & sName & "','" & sCapitalCity & "','" & sPhoneCode & "'"
        strSql = strSql & " ,'" & sCountryFlag & "','" & idCurrency & "','" & idContinent & "');"
        strSql = strSql & " SELECT LAST_INSERT_ID() AS id; "
        strSql = strSql & " END; "
        strSql = strSql & " END IF; "
        Set rs = con.Execute(strSql)
        If Not rs.EOF Then
            idISOCode = rs!id 'Guarda o Id do pais inserido ou atualizado..
        End If
 
        For Each listLanguange In countryInfo.Languages 'CountryList.lan
            idISOCodeL = 0
            sISOCodeL = listLanguange.ISOCode
            sNameL = listLanguange.Name
            'String para verificar se o indioma do pais existe, se existir carrega na variavel, se nao existir insere e pega o id
            strSql = "IF (SELECT 1 = 1 FROM tb_language WHERE language_iso_code='" & sISOCodeL & "') THEN "
            strSql = strSql & " BEGIN "
            strSql = strSql & " SELECT id FROM tb_language WHERE language_iso_code='" & sISOCodeL & "';"
            strSql = strSql & " END; "
            strSql = strSql & " ELSE "
            strSql = strSql & " BEGIN "
            strSql = strSql & " INSERT INTO tb_language(language_iso_code,name_language) "
            strSql = strSql & " VALUES('" & sISOCodeL & "','" & sNameL & "');"
            strSql = strSql & " SELECT LAST_INSERT_ID() AS id; "
            strSql = strSql & " END; "
            strSql = strSql & " END IF; "
            Set rs = con.Execute(strSql)
            If Not rs.EOF Then
                idISOCodeL = rs!id
            End If
            'String para salvar os indiomas do pais
            strSql = "IF (SELECT 1 = 1 FROM tb_country_info WHERE id_country='" & idISOCode & "' and id_language='" & idISOCodeL & "') THEN "
            strSql = strSql & " BEGIN "
            strSql = strSql & " SELECT id FROM tb_country_info WHERE id_country='" & idISOCode & "' and id_language='" & idISOCodeL & "';"
            strSql = strSql & " END; "
            strSql = strSql & " ELSE "
            strSql = strSql & " BEGIN "
            strSql = strSql & " INSERT INTO tb_country_info(id_country,id_language) "
            strSql = strSql & " VALUES('" & idISOCode & "','" & idISOCodeL & "');"
            strSql = strSql & " SELECT LAST_INSERT_ID() AS id; "
            strSql = strSql & " END; "
            strSql = strSql & " END IF; "
            Set rs = con.Execute(strSql)
            If Not rs.EOF Then
                idISOCodeL = rs!id 'Guarda o Id do pais inserido ou atualizado..
            End If
        Next listLanguange
    Next countryInfo
End Sub

Private Function ReadIniFile(ByVal strINIFile As String, ByVal strSECTION As String, ByVal strKey As String) As String
'Função para captura e leitura do arquivo INI
    Dim strBuffer As String
    Dim intPos As Integer
    strBuffer = Space$(255)
    If GetPrivateProfileString(strSECTION, strKey, "", strBuffer, 255, strINIFile) > 0 Then
        ReadIniFile = RTrim$(StripTerminator(strBuffer))
    Else
        ReadIniFile = ""
    End If
End Function


Private Function StripTerminator(ByVal strString As String) As String
'Função para captura e leitura do arquivo INI
    Dim intZeroPos As Integer

    intZeroPos = InStr(strString, Chr$(0))
    If intZeroPos > 0 Then
        StripTerminator = Left$(strString, intZeroPos - 1)
    Else
        StripTerminator = strString
    End If
End Function
