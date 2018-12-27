Private pId As String
Private pName As String
Private pHierarchyName As String
Private pLevel As Integer
Private pChildren As Collection
Private pParentId As String
Private pAdditionalFields As Collection


Private Sub Class_Initialize()
    Set pChildren = New Collection
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

Public Property Let dimensionName(newName As String)
    pDimensionName = newName
End Property

Public Property Get Children() As Collection
    Set Children = pChildren
End Property

Public Property Let Children(newChildren As Collection)
    Set pChildren = newChildren
End Property


Public Property Get IsLeaf() As Boolean
   If pChildren.Count = 0 Then
       IsLeaf = True
    Else
        IsLeaf = False
    End If
End Property

Public Property Get IsRoot() As Boolean
    If pParentId = "" Then
        IsRoot = True
    Else
        IsRoot = False
    End If
End Property

Public Property Get ParentId() As String
    ParentId = pParentId
End Property

Public Property Let ParentId(newParentId As String)
    pParentId = newParentId
End Property


Public Sub AddChild(child As HierarchyMember)
    pChildren.Add child
End Sub


Public Property Let Level(newLevel As Integer)
     pLevel = newLevel
End Property

Public Property Get Level() As Integer
    Level = pLevel
End Property


Public Property Let HierarchyName(newHierarchyName As String)
    pHierarchyName = newHierarchyName
End Property

Public Property Get HierarchyName() As String
    If pHierarchyName = "" Then
        Dim tabs As String
        For i = 1 To pLevel - 1
            tabs = tabs & vbTab
        Next i
        HierarchyName = CStr(tabs & pName)
    Else
        HierarchyName = pHierarchyName
    End If
End Property


Public Sub AddField(field As AdditionalField)
    pAdditionalFields.Add field
End Sub


Public Property Get additionalFields() As Collection
    Set additionalFields = pAdditionalFields
End Property







