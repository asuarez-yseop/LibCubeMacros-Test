'Doc: This module creates XML cube dimensions using the data access functions as data source
Public Sub createDimensions()
    createDimTime
    createDimCity
    createDimProduct
    createDimStore
End Sub


Private Sub createDimTime()
    Dim timeMembers As Collection
    
    Set timeMembers = getTimeMembers()
    
    createBasicTimeDimension timeMembers
    
    createElement element:="currentPeriod", _
                  parent:=yInstElem, _
                  YID:="TIME_YEAR_2018"
    
    createElement element:="previousPeriod", _
                  parent:=yInstElem, _
                  YID:="TIME_YEAR_2017"
                  
End Sub
Private Sub createDimCity()
    Dim cities As Collection
    
    Set cities = getCities()
    createBasicHierarchicalDim "City", cities
    
End Sub


Private Sub createDimProduct()
   Dim rootProduct As HierarchyMember
    
   Set rootProduct = getProductHierarchy()
   createHierarchicalDimension rootProduct, "Product"
End Sub



Private Sub createDimStore()
   Dim stores As Collection
    
   Set stores = getStores()
   createBasicHierarchicalDim "Store", stores
End Sub


