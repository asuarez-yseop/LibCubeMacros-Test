Private pId As String
Private pName As String
Private pAdditionalFields As Collection


Private Sub Class_Initialize()
    Set pAdditionalFields = New Collection
End Sub


Public Property Get id() As String
    id = pId
End Property

Public Property Let id(newId As String)
    pId = newId
End Property

Public Property Get name() As String
    name = pName
End Property

Public Property Let name(newName As String)
    pName = newName
End Property


Public Sub AddField(field As AdditionalField)
    pAdditionalFields.Add field
End Sub


Public Property Get additionalFields() As Collection
    Set additionalFields = pAdditionalFields
End Property

