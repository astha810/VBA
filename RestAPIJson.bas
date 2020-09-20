Attribute VB_Name = "RestAPIJson"
Option Explicit

Sub jsonimp()
    Dim strJson As String
    Dim jp As Object
    Dim dict
    Dim subdict
    Dim filepath As String
    Dim sfile As Integer
    Dim strline As String
    Dim req As WinHttpRequest
    Dim url As String
    Dim rolejson
    Dim jproles As Object
    Dim roles
    Dim role
    Dim attr
    Dim usid
    
    
    
    
    ' Reading from file
    filepath = "C:\Users\cgupta\Desktop\astha\MS Excel" + "\" + "auth.json"
    sfile = FreeFile
    strJson = ""
    Open filepath For Input As sfile
    Do While Not EOF(sfile)
        Line Input #sfile, strline
        strJson = strJson + strline
    Loop
    
    Close sfile
    
    
    
    Set jp = JsonConverter.ParseJson(strJson)
    Set req = New WinHttpRequest
    url = "https://api.harvestapp.com/v2/roles"
    req.Open "GET", url
    
    
    For Each dict In jp
    
    
        Debug.Print dict
        Debug.Print jp(dict)
        req.SetRequestHeader dict, jp(dict)
        
        
        
   
    
    Next dict
    req.Send
    rolejson = req.ResponseText
    Debug.Print rolejson
    Set jproles = JsonConverter.ParseJson(rolejson)
    Set roles = jproles("roles")
    For Each role In roles
        For Each attr In role
            Debug.Print attr
            If Not IsObject(role(attr)) Then
                Debug.Print role(attr)
              
            Else
                For Each usid In role(attr)
                    Debug.Print (usid)
                    
                Next usid
            End If
        Next attr
        
        Debug.Print vbNewLine
        
    
    Next role
    
    
End Sub


