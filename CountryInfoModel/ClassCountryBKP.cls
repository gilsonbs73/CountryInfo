VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClassCountryBKP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function GetCountryIsoCode(URI As String, FilteredIsoCode As String)
    Dim resultSet As New DOMDocument60
    Dim countryModel As New CountryInfoModel.FullCountryInfo
    Dim APITextResponse As String

    If URI = "" Then URI = "http://webservices.oorsprong.org/websamples.countryinfo/CountryInfoService.wso/FullCountryInfoAllCountries"
          
    APITextResponse = countryModel.GetFullCountryInfo(URI)
    
    resultSet.loadXML (APITextResponse)
    
    'Var xmlFiltered = resultSet.getElementsByTagName("tCountryInfo")
    
    'Filtrar os paises com a ISO Code A
    
     Dim doc As New MSXML2.DOMDocument
    Dim success As Boolean

   success = doc.loadXML(rText.Text)
   'success = doc.Load("c:\tmp\txt.xml")
   
   If success = False Then
      MsgBox doc.parseError.reason
   Else
      Dim nodeCountry As MSXML2.IXMLDOMNodeList
      Dim nodeLanguage As MSXML2.IXMLDOMNodeList



      Set nodeCountry = doc.selectNodes("/ArrayOftCountryInfo/tCountryInfo")
      'Lista os paisas
      If Not nodeCountry Is Nothing Then
         Dim nodeC As MSXML2.IXMLDOMNode
         Dim nodeL As MSXML2.IXMLDOMNode
         'Variaveis do pais
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
         

         For Each nodeC In nodeCountry
            ' Could also do node.attributes.getNamedItem("name").text
            sISOCode = nodeC.selectSingleNode("sISOCode").Text
            If Left(sISOCode, 1) = txtFiltro.Text Then
                idISOCode = 0
                idCurrency = 0
                idCurrency = 0
                
                sName = Replace(nodeC.selectSingleNode("sName").Text, "'", " ")
                sCapitalCity = Replace(nodeC.selectSingleNode("sCapitalCity").Text, "'", " ")
                sPhoneCode = nodeC.selectSingleNode("sPhoneCode").Text
                sContinentCode = nodeC.selectSingleNode("sContinentCode").Text
                sCurrencyISOCode = nodeC.selectSingleNode("sCurrencyISOCode").Text
                sCountryFlag = nodeC.selectSingleNode("sCountryFlag").Text
               
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
                '
                'Filtra as linguagens de todos os paises
                For Each nodeL In nodeC.selectSingleNode("Languages").childNodes
                    idISOCodeL = 0
                    sISOCodeL = nodeL.selectSingleNode("sISOCode").Text
                    sNameL = Replace(nodeL.selectSingleNode("sName").Text, "'", " ")
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
                Next nodeL
            End If
         Next nodeC
      End If
   End If
    
    'GetCountryIsoCode = resultSet.Text
     
     GetCountryIsoCode = APITextResponse
     DoEvents
End Function




