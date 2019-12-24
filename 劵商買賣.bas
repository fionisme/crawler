'Request URL: https://tw.stock.yahoo.com/d/s/major_2330.html
'Request Method: GET
Sub Yahoomajor()

'1help
Dim oXML As Object
Set oXML = CreateObject("winhttp.winhttprequest.5.1")

'2 send Request
'3Response
With oXML
    .Open "GET", "https://tw.stock.yahoo.com/d/s/major_2330.html", 0
    .send
    Debug.Print .responseText
    
'4go home
Set oXML = Nothing



End With

End Sub



