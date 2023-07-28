Attribute VB_Name = "Ä£¿é1"
Function GetTruckDistance(origin As String, destination As String, height As Double, _
width As Double, weight As Double, length As Double, axle_count As Integer, is_trailer As Integer, _
plate_province As String, plate_number As String, plate_color As Integer, power_type As Integer, _
truck_type As Integer, emission_limit As Integer, load_weight As Double) As Double
    Dim XMLHTTP As Object
    Dim baiduUrl As String
    Dim API_Key As String
    Dim JSON As Object
    Dim jsonResponse As Object
    Dim distance As Double

    Set XMLHTTP = CreateObject("MSXML2.XMLHTTP")
    API_Key = "Bo1FXfXT0xNH7OqKYG2aN9UtBUKrQ1Rv" ' Replace this with your Baidu Map API key
    baiduUrl = "http://api.map.baidu.com/directionlite/v1/driving?origin=" & origin & "&destination=" & destination _
    & "&heigh=" & height & "&width=" & width & "&weight=" & weight & "&length=" & lenghth & "&axle_count=" & axle_count & "&axle_count=" & axle_count _
    & "&is_trailer=" & is_trailer & "&plate_province=" & plate_province & "&plate_number=" & plate_number & "&plate_color=" & plate_color & "&power_type=" & power_type _
     & "&truck_type=" & truck_type & "&emission_limit=" & emission_limit & "&load_weight=" & load_weight & "&ak=" & API_Key

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
    
    GetTruckDistance = distance / 1000
    
End Function


