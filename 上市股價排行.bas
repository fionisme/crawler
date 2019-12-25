Attribute VB_Name = "Module1"
'Request URL: https://tw.stock.yahoo.com/d/i/rank.php?t=pri&e=tse
'Request Method: GET
Sub 上市股價排行表()
Dim oXML  As Object
Set oXML = CreateObject("winhttp.winhttprequest.5.1")


With oXML
    .Open "GET", "https://tw.stock.yahoo.com/d/i/rank.php?t=pri&e=tse", 0
    .send
    Debug.Print .responseText
End With
Set oXML = Nothing


End Sub
