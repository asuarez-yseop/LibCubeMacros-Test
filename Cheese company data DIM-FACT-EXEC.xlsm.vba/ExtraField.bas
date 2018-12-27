Private pKbFieldName As String
Private pId As String
Private pIsVariable As Boolean
Private pValue As String
Private pLabel As String
Private pVariableType As String
Private pMin As String
Private pMax As String
Private pStep As String


Public Property Get kbFieldName() As String
    kbFieldName = pKbFieldName
End Property

Public Property Let kbFieldName(newKbFieldName As String)
    pKbFieldName = newKbFieldName
End Property

Public Property Get id() As String
    id = pId
End Property

Public Property Let id(newId As String)
    pId = newId
End Property

Public Property Get isVariable() As Boolean
    isVariable = pIsVariable
End Property

Public Property Let isVariable(isVariable As Boolean)
    pIsVariable = isVariable
End Property

Public Property Get value() As String
    value = pValue
End Property

Public Property Let value(newValue As String)
    pValue = newValue
End Property

Public Property Get label() As String
    label = pLabel
End Property

Public Property Let label(newlabel As String)
    pLabel = newlabel
End Property

Public Property Get variableType() As String
    variableType = pVariableType
End Property

Public Property Let variableType(newVariableType As String)
    pVariableType = newVariableType
End Property

Public Property Get min() As String
    min = pMin
End Property

Public Property Let min(newMin As String)
    pMin = newMin
End Property

Public Property Get max() As String
    max = pMax
End Property

Public Property Let max(newMax As String)
    pMax = newMax
End Property

Public Property Get step() As String
    step = pStep
End Property

Public Property Let step(newStep As String)
    pStep = newStep
End Property


