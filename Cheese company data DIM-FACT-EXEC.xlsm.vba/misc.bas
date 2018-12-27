' below are functions which may be useful when creating the XML file from Excel

Function toYID(theValue) As String
    
    toYID = Replace(UCase(theValue), " ", "_")
    toYID = Replace(toYID, ",", "_")
    toYID = Replace(toYID, "&", "_")
    toYID = Replace(toYID, "%", "_")
    toYID = Replace(toYID, "(", "")
    toYID = Replace(toYID, "'", "")
    toYID = Replace(toYID, "?", "")
    toYID = Replace(toYID, "í", "i")
    toYID = Replace(toYID, "Í", "I")
    toYID = Replace(toYID, ")", "")
    toYID = Replace(toYID, "-", "_")
    toYID = Replace(toYID, "/", "_")
    toYID = Replace(toYID, ".", "_")
    
End Function
Function lastRow(ws As Worksheet, Optional col As Long) As Long
Dim lr As Long
    lastRow = 1
    With ws
        If col = 0 Then
            lr = .Cells.Find("*", .Range("A1"), xlFormulas, xlPart, xlByRows, xlPrevious, 0).Row
        Else
            lr = .Cells(.rows.Count, col).End(xlUp).Row
        End If
        lastRow = lr
    End With
    Exit Function
errExit:
    ' probably empty sheet
End Function
Function ColumnRange(ws As Worksheet, col As Long, Optional includeHeader As Boolean = True) As Range
    Dim lastRowIndex As Long
    Dim theRange As Range
    lastRowIndex = lastRow(ws, CLng(col))
    
    If includeHeader = True Then
        Set theRange = ws.Range(ws.Cells(1, col), ws.Cells(lastRowIndex, col))
    Else
        Set theRange = ws.Range(ws.Cells(2, col), ws.Cells(lastRowIndex, col))
    End If
    
    Set ColumnRange = theRange
End Function

Public Function lastColumn(ws As Worksheet, rowIndex As Long) As Long
    Dim last As Long
     last = ws.Cells(rowIndex, ws.columns.Count).End(xlToLeft).column
    lastColumn = last
End Function
Public Sub DeleteEmptyCells(aRange As Range)
    Dim rng As Range

    'Store blank cells inside a variable
    On Error GoTo NoBlanksFound
    Set rng = aRange.SpecialCells(xlCellTypeBlanks)
    On Error GoTo 0
 
    'Delete blank cells and shift upward
     rng.rows.Delete Shift:=xlShiftUp

     Exit Sub

     'ERROR HANLDER
NoBlanksFound:
  Exit Sub

End Sub
Function URLencode(EncodeStr As String) As String
    Dim i As Integer
    Dim erg As String
    erg = EncodeStr

    ' First replace '%' chr
    erg = Replace(erg, "%", Chr(1))

    ' then '+' chr
    erg = Replace(erg, "+", Chr(2))
    
    For i = 0 To 255
        Select Case i
            ' Allowed 'regular' characters
            Case 37, 43, 48 To 57, 65 To 90, 97 To 122

            Case 1  ' Replace original %
                erg = Replace(erg, Chr(i), "%25")

            Case 2  ' Replace original +
                erg = Replace(erg, Chr(i), "%2B")

            Case 32 ' Replace original space
                erg = Replace(erg, Chr(i), "+")

            Case 3 To 15
                erg = Replace(erg, Chr(i), "%0" & Hex(i))

            Case Else
                erg = Replace(erg, Chr(i), "%" & Hex(i))

        End Select
    Next
    
    URLencode = erg
    
End Function

Function Base64EncodedCreds() As String
    
  Dim arrData() As Byte
  arrData = StrConv(login & ":" & password, vbFromUnicode)

  Dim objXML As DOMDocument60
  Dim objNode As IXMLDOMElement

  Set objXML = New DOMDocument60
  Set objNode = objXML.createElement("b64")

  objNode.DataType = "bin.base64"
  objNode.nodeTypedValue = arrData
  Base64EncodedCreds = objNode.text

  Set objNode = Nothing
  Set objXML = Nothing
    
End Function
Function getTextFromFile(filename As String) As String
    Dim strFilename As String: strFilename = filename
    Dim strFileContent As String
    Dim iFile As Integer: iFile = FreeFile
    Open strFilename For Input As #iFile
    strFileContent = Input(LOF(iFile), iFile)
    Close #iFile
    
    getTextFromFile = strFileContent
    
End Function
Function isFileOpen(filename As String)
    Dim filenum As Integer, errnum As Integer

    On Error Resume Next   ' Turn error checking off.
    filenum = FreeFile()   ' Get a free file number.
    ' Attempt to open the file and lock it.
    Open filename For Input Lock Read As #filenum
    Close filenum          ' Close the file.
    errnum = Err           ' Save the error number that occurred.
    On Error GoTo 0        ' Turn error checking back on.

    ' Check to see which error occurred.
    Select Case errnum

        ' No error occurred.
        ' File is NOT already open by another user.
        Case 0
         isFileOpen = False

        ' Error number for "Permission Denied."
        ' File is already opened by another user.
        Case 70
            isFileOpen = True

        ' Another error occurred.
        Case Else
            Error errnum
    End Select

End Function

Sub copyHTMLresponse()
    Set IEobject = New InternetExplorerMedium ' creates IE object
    
    With IEobject
        On Error Resume Next
            .navigate outputHTML ' tries to open result html in IE object...
            Select Case Err ' in case of errors...
                Case 0 ' no error: outputHTML exists
                    .Visible = False ' ...in the background
                    '.Visible = True ' ...with window visible
                    .ExecWB 17, 0 ' selects all...
                    .ExecWB 12, 2 ' ...and copies selection to clipboard
                Case -2147023673 ' outputHTML does not exist
                    ' do nothing - error message should popup from IE
                Case Else ' all other errors
                    MsgBox "Excel responded with: Error " & Err
            End Select
            
            .Quit ' quit IE object
        On Error GoTo 0
    End With
End Sub

Sub applyHTMLStyle(sheet As Worksheet, rangeWithText As Range)
    On Error GoTo ErrorHandler
        Dim Ie As Object
    
        Set Ie = CreateObject("InternetExplorer.Application")
    
        With Ie
            .Visible = False
    
            .navigate "about:blank"
    
            .Document.body.innerHTML = rangeWithText.value
    
            .ExecWB 17, 0
            'Select all contents in browser
            .ExecWB 12, 2
            'Copy them
            
            sheet.Paste Destination:=rangeWithText
    
            .Quit
        End With
    Exit Sub
    
ErrorHandler:
    Resume
    'I.e. re-run the line of code that caused the error
Exit Sub
    
End Sub

Sub setSheet(Optional theSheet As String)
    If theSheet = "" Then
        Set dataSheet = ActiveWorkbook.ActiveSheet
        dataSheet.Select
        dataSheet.Activate
    Else
        Set dataSheet = ActiveWorkbook.Sheets(theSheet)
        dataSheet.Select
        dataSheet.Activate
    End If
End Sub

Function Find_First(theString As String) As Range
    Dim FindString As String
    Dim rng As Range
    FindString = theString
    If Trim(FindString) <> "" Then
        With ActiveSheet.Range("A1:" & Range("A1").SpecialCells(xlCellTypeLastCell).Address)
            Set rng = .Find(What:=FindString, _
                            After:=.Cells(.Cells.Count), _
                            LookIn:=xlValues, _
                            LookAt:=xlWhole, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not rng Is Nothing Then
                Set Find_First = rng
            Else
                MsgBox "Nothing found"
            End If
        End With
    End If
End Function


Function FindColumn(ws As Worksheet, theString As String, Optional rowIndex As Long = 1) As Long
    Dim lastCol As Long
    Dim columnsRange As Range
    Dim resultRange As Range
    
    lastCol = lastColumn(ws, rowIndex)
    Set columnsRange = ws.Range(ws.Cells(rowIndex, 1), ws.Cells(rowIndex, lastCol))
    Set resultRange = columnsRange.Find(theString, , xlValues, xlWhole)
    If Not resultRange Is Nothing Then
        FindColumn = resultRange.column
        Exit Function
    End If
    
    FindColumn = -1
    
End Function


Function Find_Row(theString As String) As Integer
    Dim FindString As String
    Dim rng As Range
    FindString = theString
    Dim asdf As Worksheet
    Set asdf = Sheets("Transaction")

    If Trim(FindString) <> "" Then
        With asdf.Range("A1:A100")
            Set rng = .Find(What:=FindString, _
                            After:=asdf.Cells(1), _
                            LookIn:=xlValues, _
                            LookAt:=xlPart, _
                            SearchOrder:=xlByRows, _
                            SearchDirection:=xlNext, _
                            MatchCase:=False)
            If Not rng Is Nothing Then
                Find_Row = rng.Row
            Else
                MsgBox "Nothing found"
            End If
        End With
    End If
End Function

Sub DeleteTableRows(ByRef table As ListObject)
    On Error Resume Next
    '~~> Clear Header Row `IF` it exists
    table.DataBodyRange.rows(1).ClearContents
    '~~> Delete all the other rows `IF `they exist
    table.DataBodyRange.Offset(1, 0).Resize(table.DataBodyRange.rows.Count - 1, _
    table.DataBodyRange.columns.Count).rows.Delete
    On Error GoTo 0
End Sub
Public Sub deleteSheetIfExists(sheetName As String)
    If SheetExists(sheetName) = True Then
        Application.DisplayAlerts = False
        Worksheets(sheetName).Delete
        Application.DisplayAlerts = True
    End If
End Sub
Function SheetExists(shtName As String, Optional wb As Workbook) As Boolean
    Dim sht As Worksheet

     If wb Is Nothing Then Set wb = ThisWorkbook
     On Error Resume Next
     Set sht = wb.Sheets(shtName)
     On Error GoTo 0
     SheetExists = Not sht Is Nothing
 End Function

Public Sub recreateSheet(sheetName As String)
    Application.DisplayAlerts = False
    On Error Resume Next
    Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    Worksheets.Add.name = sheetName
End Sub


Public Function getSelectedFilterValues(pTable As pivotTable, filterName As String) As Collection
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim selectedValues As Collection
    
    Set selectedValues = New Collection
    pTable.PivotCache.MissingItemsLimit = xlMissingItemsNone
    'inputTable.PivotCache.Refresh

    Set pf = pTable.PivotFields(filterName)
    For Each pi In pf.PivotItems

        If pi.Visible Then
               selectedValues.Add pi.value
        End If
    Next
    
    Set getSelectedFilterValues = selectedValues
End Function

Public Function attributeCase(theString As String)
   attributeCase = LCase(Left(theString, 1)) & Mid(theString, 2)
End Function

Sub copyHTMLtoTextBox()
 
    Set IEobject = New InternetExplorerMedium ' creates IE object
 
    With IEobject
        On Error Resume Next
            .navigate outputHTML ' tries to open result html in IE object...
            Select Case Err ' in case of errors...
                Case 0 ' no error: outputHTML exists
                    .Visible = False ' ...in the background
                    .ExecWB 17, 0 ' selects all...
                    .ExecWB 12, 2 ' ...and copies selection to clipboard
                Case -2147023673 ' outputHTML does not exist
                    ' do nothing - error message should popup from IE
                Case Else ' all other errors
                    MsgBox "Excel responded with: Error " & Err
            End Select
 
            .Quit ' quit IE object
        On Error GoTo 0
    End With
    
    ' find the row of the current data
    Set dataSheet = ActiveWorkbook.Worksheets("COST cube 2")
    dataSheet.Activate
    Dim dataRow As Integer
    dataRow = 8
    
     ' deletes shape "result" if it exists
    On Error Resume Next
        ActiveSheet.Shapes("result").Delete
    On Error GoTo 0
    
    ' creates a text box
    With Range("V" & dataRow & ":Z" & dataRow + 20)
        l = .Left
        t = .Top
        w = .Width
        h = .Height
    End With
    ActiveSheet.Shapes.AddTextbox(msoTextOrientationHorizontal, l, t, w, h).name = "result"
    
    ' pastes the result into the text box
    ActiveSheet.Shapes.Range(Array("result")).TextFrame2.TextRange.Paste
 
 
End Sub