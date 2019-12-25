Attribute VB_Name = "Module4"
Sub test()

    Dim rawResponseText As String
    Dim oXML As Object
    Set oXML = CreateObject("WinHttp.WinHttpRequest.5.1")
    Dim stock As String
    stock = InputBox("請輸入股票代號")
    
    
    With oXML
        .Open "POST", "https://mops.twse.com.tw/mops/web/ajax_t164sb04", 0
        .setRequestHeader "accept", "*/*"
        .setRequestHeader "accept-language", "zh-TW,zh;q=0.9,en-US;q=0.8,en;q=0.7"
        .setRequestHeader "cookie", "annual2019-promotion=1; annual2019-static-promotion=1; _fbc=fb.1.1575250988296.IwAR3mANDaF4a2RUXMnWsXkMkZGuDSThP_zx-_zyjRLQOdUEqaX6EhyyhiIhI; _fbp=fb.1.1575250988298.1785" & _
            "051480; BrowserMode=Web; _hjid=0bfe0bae-e0ff-4c1f-be70-cfa04d0a28eb; _ga=GA1.2.1469980782.1575250989; _gid=GA1.2.1288133104.1575250989; _gat_gtag_UA_6993262_2=1; wcsid=c6" & _
            "f4edZ49uT2asM53h7B70HB6a3G6obk; hblid=PYXjeyTV5Ln1wFx33h7B70HB38r6bafS; _oklv=1575250988853%2Cc6f4edZ49uT2asM53h7B70HB6a3G6obk; _okdetect=%7B%22token%22%3A%22157525098916" & _
            "70%22%2C%22proto%22%3A%22https%3A%22%2C%22host%22%3A%22www.wantgoo.com%22%7D; olfsk=olfsk8520169367067283; _ok=8391-691-10-7433; _okbk=cd5%3Davailable%2Ccd4%3Dtrue%2Cvi5%" & _
            "3D0%2Cvi4%3D1575250991796%2Cvi3%3Dactive%2Cvi2%3Dfalse%2Cvi1%3Dfalse%2Ccd8%3Dchat%2Ccd6%3D0%2Ccd3%3Dfalse%2Ccd2%3D0%2Ccd1%3D0%2C; cf_clearance=06c7fa7b4c6903cbdac93d6d382" & _
            "e9868a00130ab-1575251000-0-150; __cfduid=daafadbe7633cb91556b8d4e1d1bf13381575251000; BID=76A62CD2-A79D-4E28-B070-A1F3044A441C"
        .setRequestHeader "referer", "https://www.wantgoo.com/stock/astock/techchart?stockno=2330&fbclid=IwAR3mANDaF4a2RUXMnWsXkMkZGuDSThP_zx-_zyjRLQOdUEqaX6EhyyhiIhI&__cf_chl_jschl_tk__=a2f8100c5ab0a2c863bd5" & _
            "28f9f3b025000d29d11-1575250984-0-AWV60Vq5S1lEBfMJMG55osxC52QKoamxdKdRj2cEN2-SzET0My9oM3gOW6ddqRPjknDR13kbhSjOUO4wH1MkRIG01Ch_Ta5BNl1chfadefQGetUFJ_L3Ms4QbnNnuPTEWs7eVPeab" & _
            "KjCsLVE3aKHmKfaDd7b5DdSQGF-S4r9AcHBvP0E-mAX1BoCPJXLoHJcQUtu1DNa5PLfRrYmiG0w_hcDbk2H6hj3UDzU5bK6f_juT0RsUojnugWRHYqFZJJBSkxanMDIQikEIW0L1cxAFLVqU5oJixodh6uiRQpcYUGjOdrxdFB" & _
            "zOw995IdtRJ25IozHejThebenDiFqRNGrG3d1q56tJR_mPKCVgEEHGKkCGraZlNl5fgLtOqoOHOe4TpnWSbzLmE5KRV1zwC6KrkGuOZvatvur5oPi9F22b0arxr5npk7WNivh_knGG1LG5lgLBeewKkWxZfOqh8Aze6s"
        .setRequestHeader "sec-fetch-mode", "cors"
        .setRequestHeader "sec-fetch-site", "same-origin"
        .setRequestHeader "user-agent", "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/78.0.3904.108 Safari/537.36"
        .setRequestHeader "x-requested-with", "XMLHttpRequest"
        .send "encodeURIComponent=1&step=1&firstin=1&off=1&keyword4=&code1=&TYPEK2=&checkbtn=&queryName=co_id&inpuType=co_id&TYPEK=all&isnew=true&co_id=3008&year=&season="

        rawResponseText = convertraw(.responseBody, "UTF-8")
        Debug.Print rawResponseText
    End With
    Set oXML = Nothing
End Sub

Function convertraw(rawdata, char)

Dim rawstr
Set rawstr = CreateObject("adodb.stream")
With rawstr
  .Type = 1
  .Mode = 3
  .Open
  .Write rawdata
  .Position = 0
  .Type = 2
  .Charset = char
  convertraw = .ReadText
  .Close
End With
Set rawstr = Nothing

End Function



