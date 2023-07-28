Attribute VB_Name = "ģ��6"
Function GetLatLng(Address As String) As String

    Dim XMLHTTP As Object, baiduUrl As String, API_Key As String
    Dim JSON As Object

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    API_Key = "your api key" ' replace this with your Baidu Map API key
    baiduUrl = "http://api.map.baidu.com/geocoding/v3/?address=" & Address & "&output=json&ak=" & API_Key

    XMLHTTP.Open "GET", baiduUrl, False
    XMLHTTP.setRequestHeader "Content-Type", "application/json"
    XMLHTTP.send
    
    ' Parse the response
    Set JSON = ParseJson(XMLHTTP.responseText)
    GetLatLng = JSON("result")("location")("lat") & "," & JSON("result")("location")("lng")

End Function
