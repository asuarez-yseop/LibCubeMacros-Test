Private pId As String
Private pName As String


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

