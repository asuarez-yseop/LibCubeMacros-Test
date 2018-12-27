Private pMeasure As String
Private pValue As Double
Private pTextValue As String
Private pMeasureInfo As MeasureInfo


Public Property Get MeasureInfo() As MeasureInfo
    Set MeasureInfo = pMeasureInfo
End Property

Public Property Let MeasureInfo(newMeasureInfo As MeasureInfo)
    Set pMeasureInfo = newMeasureInfo
End Property

Public Property Get Measure() As String
    Measure = pMeasure
End Property

Public Property Let Measure(newMeasure As String)
    pMeasure = newMeasure
End Property


Public Property Get value() As Double
    value = pValue
End Property

Public Property Let value(newValue As Double)
    pValue = newValue
End Property

Public Property Get textValue() As String
    textValue = pTextValue
End Property

Public Property Let textValue(newTextValue As String)
    pTextValue = newTextValue
End Property

