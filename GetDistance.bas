Attribute VB_Name = "ģ��5"
Function GetDistance(Origin As String, Destination As String) As Double
    Dim XMLHTTP As Object
    Dim baiduUrl As String
    Dim API_Key As String
    Dim JSON As Object
    Dim jsonResponse As Object
    Dim distance As Double

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    API_Key = "your apikey" ' Replace this with your Baidu Map API key
    baiduUrl = "http://api.map.baidu.com/directionlite/v1/driving?origin=" & Origin & "&destination=" & Destination & "&ak=" & API_Key

    XMLHTTP.Open "GET", baiduUrl, False
    XMLHTTP.setRequestHeader "Content-Type", "application/json"
    XMLHTTP.send
    
    ' Print the response to the Immediate Window
    Debug.Print XMLHTTP.responseText
    
    Set JSON = JsonConverter.ParseJson(XMLHTTP.responseText)

    ' Check if the "result" property exists in the response
    If JSON("result").Count > 0 Then
        ' Get the routes array
        Set jsonResponse = JSON("result")
        If jsonResponse("routes").Count > 0 Then
            ' Get the routes
            For Each route In jsonResponse("routes")
                distance = distance + CDbl(route("distance"))
            Next route
        End If
    End If
    
    GetDistance = distance / 1000
    
End Function

