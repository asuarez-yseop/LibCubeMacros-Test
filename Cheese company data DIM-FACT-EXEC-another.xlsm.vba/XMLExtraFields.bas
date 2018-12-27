'Doc: This module creates XML extra fields using the data access functions as data source
Public Sub createExtraFields()
    Dim extraFields As Collection
    Dim fieldExtraProps As Scripting.Dictionary
    
    Set extraFields = getExtraFields()
    
    For Each field In extraFields
        If field.id <> "" Then
              createElement element:=field.kbFieldName, _
                  parent:=yInstElem, _
                  YID:=field.id
        ElseIf field.value <> "" Then
            Set fieldExtraProps = New Scripting.Dictionary
              
            fieldExtraProps.Add "w:variable", LCase(CStr(field.isVariable))
            fieldExtraProps.Add "w:label", field.label
            fieldExtraProps.Add "w:type", field.variableType
            fieldExtraProps.Add "w:min", field.min
            fieldExtraProps.Add "w:max", field.max
            fieldExtraProps.Add "w:step", field.step
              
            createElement element:=field.kbFieldName, _
                parent:=yInstElem, _
                extraProps:=fieldExtraProps, _
                theText:=field.value
        End If
    Next field
    
End Sub
