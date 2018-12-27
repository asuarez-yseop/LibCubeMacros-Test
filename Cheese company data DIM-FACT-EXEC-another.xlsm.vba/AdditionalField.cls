'Doc: Class that represents an additional field to a Cube Member

Private pColumnName As String
Private pValue As String
Private pTagName As String
Private pIsYClassField As Boolean


Private Sub Class_Initialize()
    pIsYClassField = False
End Sub

Public Property Get columnName() As String
    columnName = pColumnName
End Property

Public Property Let columnName(newColumnName As String)
    pColumnName = newColumnName
End Property

Public Property Get value() As String
    value = pValue
End Property

Public Property Let value(newValue As String)
    pValue = newValue
End Property

Public Property Get tagName() As String
    tagName = pTagName
End Property

Public Property Let tagName(newTagName As String)
    pTagName = newTagName
End Property


Public Property Get isYClassField() As Boolean
    isYClassField = pIsYClassField
End Property

Public Property Let isYClassField(isYClass As Boolean)
    pIsYClassField = isYClass
End Property





