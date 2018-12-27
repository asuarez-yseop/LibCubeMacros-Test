Public Sub createFacts()
    createSalesFacts
End Sub

Private Sub createSalesFacts()
    Dim facts As Collection
    
    Set facts = getFacts()
    
    createFactsFromDicts facts
    
    Set facts = Nothing
End Sub
