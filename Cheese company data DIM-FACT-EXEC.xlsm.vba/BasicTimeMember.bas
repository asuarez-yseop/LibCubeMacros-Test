Private pId As String
Private pPeriod As String
Private pDate As String


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





