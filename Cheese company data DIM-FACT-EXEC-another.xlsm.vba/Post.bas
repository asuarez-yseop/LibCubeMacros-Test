Sub PostYseop()
 
    Dim inputString As String ' creates a variable to store the inputXML
    Dim fileNumber As Integer ' creates a file to read the inputXML
    fileNumber = FreeFile()
 
    Open inputXML For Input As #fileNumber ' opens the input file
        Do Until EOF(fileNumber) ' goes through the input file...
            Line Input #fileNumber, textline ' ... line by line...
            inputString = inputString & textline ' to create a string (to URL encode)
        Loop
    Close #fileNumber 'closes the file
 
    Dim objHTTP As New MSXML2.XMLHTTP60 ' creates a HTTP request object
    
    Debug.Print "HTTP Obj created"
    ' construct the URL for the POST request
    postURL = baseURL & "/yseop-manager/direct/" & projectName & _
              "/dialog.do?transformation=" & transform
    Debug.Print "postURL: " & postURL
    With objHTTP
        .Open "POST", postURL, False ' sets it as a POST request
        ' request headers
        .setRequestHeader "Content-type", _
                          "application/x-www-form-urlencoded;charset=UTF-8"
        ' authorization headers
        .setRequestHeader "Authorization", _
                          "Basic " & Base64EncodedCreds()
    End With
 
    On Error Resume Next ' in case there is an error, continue to error handler
        objHTTP.send "xml=" & URLencode(inputString)
 
        ' prepare a stream file for writing
        Set objStream = CreateObject("ADODB.Stream")
        objStream.Charset = "UTF-8"
        objStream.Open
        
 
        Select Case Err ' Excel error codes
            Case 0 ' no error, server successfully contacted
                Debug.Print "objHTTP.Status: " & objHTTP.Status
                Select Case objHTTP.Status ' HTML status codes
                    Case 200 ' successful POST
                        objStream.WriteText objHTTP.responseText
 
                    Case 404 ' 404 error
                        objStream.WriteText "ERROR 404: The requested resource is not available"
 
                    Case 500 ' 500 error
                        objStream.WriteText "Text generation failed: Yseop Engine error. " & _
                                            "Please ensure your data is correct."
                    Case 503 ' 503 error
                        objStream.WriteText "ERROR 503: The application was stopped"
                    
 
                    Case statusCode ' unknown HTML status code
                        MsgBox statusCode
                        objStream.WriteText genericError & _
                                            "Server responded with: Error " & statusCode
                End Select
 
            Case -2146697211 ' server does not exist (or is offline?)
                objStream.WriteText genericError
 
            Case -2147024891 ' use HTTPS
                objStream.WriteText "Access Denied. You do not have the right privileges. " & _
                                  "Possible cause: your URL requires the use of HTTPS."
            
            Case Err ' unknown error
                objStream.WriteText genericError & _
                                    "Excel responded with: Error " & Err
 
        End Select
        Debug.Print "Err: " & Err
        'Debug.Print objHTTP.Status
        'Deubg.Print objHTTP.responseText
        'Debug.Print postURL
        
    On Error GoTo 0 ' turns off error handler
 
    objStream.SaveToFile outputHTML, 2 ' saves the file
    objStream.Close ' closes the stream
    Debug.Print "Post complete"
 
End Sub


