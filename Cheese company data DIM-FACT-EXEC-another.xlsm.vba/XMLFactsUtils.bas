'Doc: Creates the Facts XML input from a collection of facts dictionaries,
'     if the key of the dict starts with MEASURE_ then it will generate a measure tag with the dict value as a factMeasure value,
'     if that's not the case, it will generate a "member" tag with the dict value as the yid of the member

Public Sub createFactsFromDicts(factDicts As Collection)
    Dim Fact As Scripting.Dictionary
    Dim factElem As Object
    Dim factMeasureElem As Object
    
    For Each Fact In factDicts
        createElement element:="facts", _
                  parent:=cubeElem, _
                  yclass:="LibCube:Fact", _
                  setElement:=factElem
                  
        For Each key In Fact.Keys
        
            If UCase(key) Like "MEASURE_*" Then
                createElement element:="factMeasures", _
                  yclass:="LibCube:FactMeasure", _
                  parent:=factElem, _
                  setElement:=factMeasureElem
       
                createElement element:="measure", _
                           parent:=factMeasureElem, _
                           YID:=UCase(key)
                
                If IsNumeric(Fact(key)) Then
                    createElement element:="value", _
                           parent:=factMeasureElem, _
                           theText:=Fact(key)
                Else
                    If IsDate(Fact(key)) Then
                        createElement element:="textValue", _
                               parent:=factMeasureElem, _
                               theText:=Format(Fact(key), "yyyy/mm/dd")
                    Else
                        createElement element:="textValue", _
                               parent:=factMeasureElem, _
                               theText:=Fact(key)
                    End If
                End If
    
            Else
                createElement element:="members", _
                  parent:=factElem, _
                  YID:=Fact(key)
            End If
            
        Next key
    Next Fact
End Sub


Public Sub createFacts_(facts As Collection)
    Dim curFactMeasure As FactMeasure
    Dim curMeasureInfo As MeasureInfo
    Dim factElem As Object
    Dim factMeasureElem As Object
    Dim measureElem As Object
    Dim measureInfoElem As Object
    
    
    For Each Fact In facts
        createElement element:="facts", _
                      parent:=cubeElem, _
                      yclass:="LibDocument:Fact", _
                      setElement:=factElem
                      
        For Each member In Fact.members
            createElement element:="members", _
                      parent:=factElem, _
                      YID:=CStr(member)
        Next member
        
        For Each FactMeasure In Fact.factMeasures
            
            createElement element:="factMeasures", _
                      yclass:="LibCube:FactMeasure", _
                      parent:=factElem, _
                      setElement:=factMeasureElem
                      
            createElement element:="measure", _
                               parent:=factMeasureElem, _
                               YID:=UCase(FactMeasure.Measure), _
                               setElement:=measureElem
                               
            If Not FactMeasure.MeasureInfo Is Nothing Then
                createElement element:="measureInfo", _
                                   parent:=measureElem, _
                                   yclass:="LibCube:MeasureInfo", _
                                   setElement:=measureInfoElem
                
                If FactMeasure.MeasureInfo.unitType <> "" Then
                    createElement element:="unitType", _
                                   parent:=measureInfoElem, _
                                   YID:=FactMeasure.MeasureInfo.unitType
                End If
            End If
            
            If FactMeasure.textValue = "" Then
                createElement element:="value", _
                                       parent:=factMeasureElem, _
                                       theText:=FactMeasure.value
                                       
            Else
               createElement element:="textValue", _
                                       parent:=factMeasureElem, _
                                       theText:=FactMeasure.textValue
            End If
            
            
        Next FactMeasure
    Next Fact
End Sub
                                       

