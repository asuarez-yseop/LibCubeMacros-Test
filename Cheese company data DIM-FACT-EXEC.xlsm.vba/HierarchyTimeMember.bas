Private pId As String
Private pPeriod As String
Private pDate As String
Private pLevel As Integer
Private pChildren As Collection
Private pParentId As String


Public Property Get id() As String
    id = pId
End Property

Public Property Let id(newId As String)
    pId = newId
End Property

Public Property Get period() As String
    period = pPeriod
End Property

Public Property Let period(newPeriod As String)
    pPeriod = newPeriod
End Property

Public Property Get timeDate() As String
    timeDate = pDate
End Property

Public Property Let timeDate(newTimeDate As String)
    pDate = newTimeDate
End Property


'------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Methods from HierarchyMember

Private Sub Class_Initialize()
    Set pChildren = New Collection
End Sub

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


Public Sub AddChild(child As HierarchyTimeMember)
    pChildren.Add child
End Sub

Public Property Let Level(newLevel As Integer)
     pLevel = newLevel
End Property

Public Property Get Level() As Integer
    Level = pLevel
End Property