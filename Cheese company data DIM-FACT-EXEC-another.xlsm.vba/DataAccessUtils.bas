'Doc: Looksup a table using a column and a value as filter and returns the found value

Public Function lookupTableVal(sheetName As String, tableName As String, columnName As String, lookupValue As String, resultColumnName As String) As String
    Dim table As ListObject
    Dim val As String
    Dim res As String
    
    Set table = Worksheets(sheetName).ListObjects(tableName)
    
    For Each Row In table.ListRows
        val = CStr(Row.Range(table.ListColumns(columnName).Index))
        
        If val = lookupValue Then
            res = CStr(Row.Range(table.ListColumns(resultColumnName).Index))
            lookupTableVal = res
            Exit Function
        End If
        
    Next Row
    
    lookupTableVal = ""
End Function

Public Function lookupTableRow(sheetName As String, tableName As String, columnName As String, lookupValue As String) As Scripting.Dictionary
    Dim table As ListObject
    Dim val As String
    Dim resRow As Scripting.Dictionary
    
    Set table = Worksheets(sheetName).ListObjects(tableName)
    Set resRow = New Scripting.Dictionary
    
    For Each Row In table.ListRows
        val = CStr(Row.Range(table.ListColumns(columnName).Index))
        
        If val = lookupValue Then
            For Each col In table.ListColumns
               resRow.Add CStr(col), CStr(Row.Range(col.Index))
               Exit For
            Next col
        End If
        
    Next Row
    
    Set lookupTableRow = resRow
End Function


'Doc: Returns a dictionary from the data of the table name passed as parameter, with the key and value from the columns specified
Public Function getDictFromTable(sheetName As String, tableName As String, keyColumnName As String, valueColumnName As String) As Scripting.Dictionary
    Dim table As ListObject
    Dim dict As Scripting.Dictionary
    Dim val As Variant
    Dim key As Variant
    
    Set table = Worksheets(sheetName).ListObjects(tableName)
    Set dict = New Scripting.Dictionary
    
    For Each Row In table.ListRows
        key = Row.Range(table.ListColumns(keyColumnName).Index)
        val = Row.Range(table.ListColumns(valueColumnName).Index)
        
        dict.Add key, val
    Next Row
    
    Set getDictFromTable = dict
    
    
End Function

'Doc: Returns a collection of basic members from a table in a worksheet
Public Function getBasicMembersFromTable(sheetName As String, tableName As String, Optional includeAdditionalFields As Boolean = False) As Collection
    Dim table As ListObject
    Dim members As Collection
    Dim currMember As BasicMember
    Dim col As ListColumn
    Dim AdditionalField As AdditionalField
    Dim yClassName As String
    
    Set members = New Collection
    Set table = Worksheets(sheetName).ListObjects(tableName)
    
    
    For Each Row In table.ListRows
        Set currMember = New BasicMember
        
        For Each col In table.ListColumns
           If col.name = "Id" Then
              currMember.id = Row.Range(col.Index)
           ElseIf col.name = "Name" Then
              currMember.name = Row.Range(col.Index)
           Else
              If includeAdditionalFields = True Then
                Set AdditionalField = New AdditionalField
                AdditionalField.columnName = col.name
                AdditionalField.value = Row.Range(col.Index)
                
                If col.name Like "yClass:*" Then
                   yClassName = Split(col.name, ":")(1)
                   AdditionalField.tagName = attributeCase(yClassName)
                   AdditionalField.isYClassField = True
                Else
                   AdditionalField.tagName = attributeCase(col.name)
                End If
                
                currMember.AddField AdditionalField
              End If
           End If
           
        Next col
        
        members.Add currMember
    Next Row
    
    
    Set getBasicMembersFromTable = members
End Function

'Doc: Returns a collection of extra fields from a table in a worksheet
Public Function getExtraFieldsFromTable(sheetName As String, tableName As String) As Collection
    Dim table As ListObject
    Dim members As Collection
    Dim currMember As ExtraField
    Dim col As ListColumn
    
    Set members = New Collection
    Set table = Worksheets(sheetName).ListObjects(tableName)
    
    
    For Each Row In table.ListRows
        Set currMember = New ExtraField
        
        For Each col In table.ListColumns
           If col.name = "KB field name" Then
              currMember.kbFieldName = Row.Range(col.Index)
           ElseIf col.name = "Id" And Row.Range(col.Index) <> "" Then
              currMember.id = Row.Range(col.Index)
           ElseIf col.name = "label" And Row.Range(col.Index) <> "" Then
              currMember.label = Row.Range(col.Index)
           ElseIf col.name = "type" And Row.Range(col.Index) <> "" Then
              currMember.variableType = Row.Range(col.Index)
              currMember.isVariable = True
           ElseIf col.name = "min" And Row.Range(col.Index) <> "" Then
              currMember.min = Row.Range(col.Index)
           ElseIf col.name = "max" And Row.Range(col.Index) <> "" Then
              currMember.max = Row.Range(col.Index)
           ElseIf col.name = "step" And Row.Range(col.Index) <> "" Then
              currMember.step = Row.Range(col.Index)
           ElseIf col.name = "value" And Row.Range(col.Index) <> "" Then
              currMember.value = Row.Range(col.Index)
           End If
           
        Next col
        
        members.Add currMember
    Next Row
    
    
    Set getExtraFieldsFromTable = members
End Function


'Doc: Returns a collection of dictionaries containing the data from a fact table where the keys are the column names and the values the cells' data
Public Function getDictsFromFactTable(sheetName As String, tableName As String) As Collection
    Dim dicts As Collection
    Dim table As ListObject
    Dim col As ListColumn
    Dim currentDict As Scripting.Dictionary
    
    Set dicts = New Collection
    Set table = Worksheets(sheetName).ListObjects(tableName)
    
    For Each Row In table.ListRows
        Set currentDict = New Scripting.Dictionary
        
        For Each col In table.ListColumns
           currentDict.Add col.name, Row.Range(col.Index)
        Next col
        
        dicts.Add currentDict
        
    Next Row
    
    Set getDictsFromFactTable = dicts
End Function

Public Function getFactsFromTable(sheetName As String, tableName As String) As Collection
    Dim factDicts As Collection
    Dim facts As Collection
    Dim hasMeasuresAsColumns As Boolean
    
    Set facts = New Collection
    Set factDicts = getDictsFromFactTable(sheetName, tableName)
    
    'Get the first fact and see if it has measures as columns
    hasMeasuresAsColumns = True
    For Each column In factDicts.Item(1).Keys
        If UCase(column) = "MEASURE" Then
            hasMeasuresAsColumns = False
        End If
    Next column
    
    If hasMeasuresAsColumns Then
        Set facts = getFactsWithMeasuresAsColumns(factDicts)
    Else
        Set facts = getFactsWithMeasuresInRows(factDicts)
    End If
    
    Set getFactsFromTable = facts
End Function

Private Function getFactsWithMeasuresAsColumns(factDicts As Collection) As Collection
    Dim facts As Collection
    Dim currFact As Fact
    Dim currFactMeasure As FactMeasure
    Set facts = New Collection
    
    For Each factDict In factDicts
            Set currFact = New Fact
            
                For Each key In factDict.Keys
            
                    If UCase(key) Like "MEASURE_*" Then
                        Set currFactMeasure = New FactMeasure
                        currFactMeasure.Measure = UCase(key)
                        
                        If IsNumeric(factDict(key)) Then
                            currFactMeasure.value = CDbl(factDict(key))
                        Else
                            If IsDate(factDict(key)) Then
                                currFactMeasure.textValue = Format(factDict(key), "yyyy/mm/dd")
                            Else
                                currFactMeasure.textValue = CStr(factDict(key))
                            End If
                        End If
            
                    Else
                        currFact.addMember (CStr(factDict(key)))
                    End If
                
                Next key
                
                facts.Add (currFact)
            
        Next factDict
        Set getFactsWithMeasuresAsColumns = facts
End Function

Private Function getFactsWithMeasuresInRows(factDicts As Collection) As Collection
    Dim facts As Collection
    Dim currFact As Fact
    Dim currFactMeasure As FactMeasure
    Dim currMeasureInfo As MeasureInfo
    Dim factsbyKey As Scripting.Dictionary
    
    
    Set facts = New Collection
    Set factsbyKey = New Scripting.Dictionary
    
    For Each factDict In factDicts
            Set currFact = New Fact
            Set currFactMeasure = New FactMeasure
            
            For Each key In factDict.Keys
                
                If UCase(key) = "MEASURE" Then
                    currFactMeasure.Measure = CStr(factDict(key))
                
                ElseIf UCase(key) = "VALUE" Then
                    
                    If IsNumeric(factDict(key)) Then
                        currFactMeasure.value = CDbl(factDict(key))
                        
                    Else
                        If IsDate(factDict(key)) Then
                            currFactMeasure.textValue = CStr(Format(factDict(key), "yyyy/mm/dd"))
                        Else
                            currFactMeasure.textValue = CStr(factDict(key))
                        End If
                    End If
        
                ElseIf UCase(key) = "UNIT_TYPE" Then
                    Set currMeasureInfo = New MeasureInfo
                    currMeasureInfo.unitType = CStr(factDict(key))
                    currFactMeasure.MeasureInfo = currMeasureInfo
                Else
                    currFact.addMember (CStr(factDict(key)))
                End If
                
            Next key
            
            currFact.addFactMeasure currFactMeasure
            
            If factsbyKey.Exists(currFact.getMembersKey()) = True Then
               factsbyKey(currFact.getMembersKey()).addFactMeasure currFactMeasure
            Else
                factsbyKey.Add currFact.getMembersKey(), currFact
            End If
        Next factDict
        
    For Each membersKey In factsbyKey.Keys()
        facts.Add (factsbyKey(membersKey))
    Next membersKey
    
    Set factsbyKey = Nothing
    Set getFactsWithMeasuresInRows = facts
    
End Function

'Doc: Returns a list of dictionaries from the data of the table name passed as parameter, with the filter column name and the filter value of that column

Public Function getFilteredDictsFromTable(sheetName As String, tableName As String, filterColumnName As String, filterValue As String) As Collection

    Dim dicts As Collection
    Dim table As ListObject
    Dim col As ListColumn
    Dim currentDict As Scripting.Dictionary
    Dim value As Variant
    Dim comparisonValue As Variant
    
    Set dicts = New Collection
    Set table = Worksheets(sheetName).ListObjects(tableName)
    
    For Each Row In table.ListRows
        comparisonValue = Row.Range(table.ListColumns(filterColumnName).Index)
        If comparisonValue = filterValue Then
            For Each col In table.ListColumns
                value = Row.Range(col.Index)
                Set currentDict = New Scripting.Dictionary
                currentDict.Add col.name, value
                dicts.Add currentDict
            Next col
        End If
        
    Next Row
    
    Set getFilteredDictsFromTable = dicts
    
    
End Function

'Doc: Returns a collection of Basic Time Members from a table
Public Function getBasicTimeMembersFromTable(sheetName As String, tableName As String) As Collection
    Dim table As ListObject
    Dim members As Collection
    Dim currMember As BasicTimeMember
    
    Set members = New Collection
    Set table = Worksheets(sheetName).ListObjects(tableName)
    
    For Each Row In table.ListRows
        Set currMember = New BasicTimeMember
        currMember.id = Row.Range(table.ListColumns("Id").Index)
        currMember.period = Row.Range(table.ListColumns("Period").Index)
        currMember.timeDate = CStr(Row.Range(table.ListColumns("Date").Index))
        
        members.Add currMember
    Next Row
    
    
    Set getBasicTimeMembersFromTable = members
End Function


'Doc: Returns a HierarchyMember tree from a table in a worksheet
Public Function getHierarchyFromTable(sheetName As String, tableName As String, Optional hasLeveLColumn As Boolean = False, Optional includeAdditionalFields As Boolean = False) As HierarchyMember

   Dim currentMember As HierarchyMember
   Dim table As ListObject
   Dim rows As ListRows
   Dim fields As ListColumns
   Dim yClassName As String
   Dim rootMember As HierarchyMember
   Dim AdditionalField As AdditionalField
   Dim membersById As New Scripting.Dictionary
   
   Set table = Worksheets(sheetName).ListObjects(tableName)
   Set rows = table.ListRows
   Set fields = table.ListColumns
   
   
   For Each Row In rows
       Set currentMember = New HierarchyMember
       For Each col In table.ListColumns
           If col.name = "Id" Then
               currentMember.id = Row.Range(col.Index)
           ElseIf col.name = "Name" Then
               currentMember.name = Row.Range(col.Index)
           ElseIf col.name = "ParentId" Then
               currentMember.ParentId = Row.Range(col.Index)
           ElseIf col.name = "Level" Then
               If hasLeveLColumn Then
                   currentMember.Level = Row.Range(col.Index)
               End If
           Else
               If includeAdditionalFields = True Then
                   Set AdditionalField = New AdditionalField
                   AdditionalField.columnName = col.name
                   AdditionalField.value = Row.Range(col.Index)
                   
                   If col.name Like "yClass:*" Then
                      yClassName = Split(col.name, ":")(1)
                      AdditionalField.tagName = attributeCase(yClassName)
                      AdditionalField.isYClassField = True
                   Else
                      AdditionalField.tagName = attributeCase(col.name)
                   End If
                   
                   currMember.AddField AdditionalField
               End If
           End If
       Next col
       
       If currentMember.ParentId = "" Then
           Set rootMember = currentMember
       End If
       
       If Not membersById.Exists(currentMember.id) Then
           membersById.Add currentMember.id, currentMember
       End If
 
   Next Row
   
   For Each member In membersById.Items
       For Each member2 In membersById.Items
           If member2.ParentId = member.id Then
               member.AddChild (member2)
           End If
       Next member2
   Next member
   
   Set membersById = Nothing
   
   Set getHierarchyFromTable = rootMember
End Function

'Doc: Returns a HierarchyTimeMember tree from a table in a worksheet
Public Function getTimeHierarchyFromTable(sheetName As String, tableName As String) As HierarchyTimeMember

   Dim currentMember As HierarchyTimeMember
   Dim table As ListObject
   Dim rows As ListRows
   Dim fields As ListColumns
   Dim yClassName As String
   Dim rootMember As HierarchyTimeMember
   Dim membersById As New Scripting.Dictionary
   
   Set table = Worksheets(sheetName).ListObjects(tableName)
   Set rows = table.ListRows
   Set fields = table.ListColumns
   
   
   For Each Row In rows
       Set currentMember = New HierarchyTimeMember
       For Each col In table.ListColumns
           If col.name = "Id" Then
               currentMember.id = Row.Range(col.Index)
           ElseIf col.name = "Date" Then
               currentMember.timeDate = Row.Range(col.Index)
           ElseIf col.name = "Period" Then
               currentMember.period = Row.Range(col.Index)
           ElseIf col.name = "ParentId" Then
               currentMember.ParentId = Row.Range(col.Index)
           End If
       Next col
       
       If currentMember.ParentId = "" Then
           Set rootMember = currentMember
       End If
       
       If Not membersById.Exists(currentMember.id) Then
           membersById.Add currentMember.id, currentMember
       End If
 
   Next Row
   For Each member In membersById.Items
       For Each member2 In membersById.Items
           If member2.ParentId = member.id Then
               member.AddChild (member2)
           End If
       Next member2
   Next member
   
   Set membersById = Nothing
   
   Set getTimeHierarchyFromTable = rootMember
End Function