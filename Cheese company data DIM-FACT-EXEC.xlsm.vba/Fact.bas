Private pMembers As Collection
Private pFactMeasures As Collection


Private Sub Class_Initialize()
    Set pMembers = New Collection
    Set pFactMeasures = New Collection
End Sub


Public Property Get members() As Collection
    Set members = pMembers
End Property

Public Property Get factMeasures() As Collection
    Set factMeasures = pFactMeasures
End Property

Public Function getMembersKey() As String
    Dim key As String
    key = ""
    
    For Each member In pMembers
       key = key & "-" & member
    Next
    
    getMembersKey = key
End Function

Public Sub addMember(memberId As String)
    pMembers.Add memberId
End Sub

Public Sub addFactMeasure(FactMeasure As FactMeasure)
    pFactMeasures.Add FactMeasure
End Sub



