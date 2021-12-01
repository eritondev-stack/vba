Attribute VB_Name = "Api"
Sub callRestApi()

    Dim objRequest As Object
    Dim strUrl As String
    Dim blnAsync As Boolean
    Dim strResponse As String
    

    Set objRequest = CreateObject("MSXML2.XMLHTTP")
    strUrl = "https://api.covid19api.com/summary"
    blnAsync = True

    With objRequest
        .Open "GET", strUrl, blnAsync
        .SetRequestHeader "Content-Type", "application/json"
        .Send
        'spin wheels whilst waiting for response
        While objRequest.readyState <> 4
            DoEvents
        Wend
        strResponse = .responseText
    End With
    
    Set JsonResult = JsonConverter.ParseJson(strResponse)
For Each M In JsonResult("Countries")
   Debug.Print M("Country")
Next M

    'Debug.Print strResponse

End Sub
