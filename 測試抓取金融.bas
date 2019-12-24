Attribute VB_Name = "Module1"
Sub text()

Dim oXML As Object
Set oXML = CreateObject("WinHttp.WinHttpRequest.5.1")
With oXML

    
    'open Request method,Request URL,False
    'send
    oXML.send
    
    '
    '.responseText
    Debug.Print oXML.responseText
End With

Set .oXML = Nothing

End Sub