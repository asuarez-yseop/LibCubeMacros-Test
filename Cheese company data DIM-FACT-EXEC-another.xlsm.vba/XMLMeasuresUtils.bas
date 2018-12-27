Public Sub createMeasures_(measures As Collection)
    Dim measureElement As Object
    For Each Measure In measures
        createElement element:="measures", _
                  parent:=cubeElem, _
                  YID:=Measure.id, _
                  yclass:="LibCube:Measure", _
                  setElement:=measureElement
                  
        createElement element:="label", _
                  parent:=measureElement, _
                  theText:=Measure.name
                  
    Next Measure
End Sub

