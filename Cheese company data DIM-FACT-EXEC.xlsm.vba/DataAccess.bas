'Doc: this module contains functions that lookup into the fact and dimension tables and return data in memory structures that are easier to use

Public Const CITIES_SHEET_NAME As String = "Cities"
Public Const CITIES_TABLE_NAME As String = "CitiesTable"

Public Const STORES_SHEET_NAME As String = "Stores"
Public Const STORES_TABLE_NAME As String = "StoresTable"

Public Const PRODUCTS_SHEET_NAME As String = "Products"
Public Const PRODUCTS_TABLE_NAME As String = "ProductsTable"

Public Const TIME_SHEET_NAME As String = "Time"
Public Const TIME_TABLE_NAME As String = "TimeTable"

Public Const SALES_FACTS_SHEET_NAME As String = "Sales"
Public Const SALES_FACTS_TABLE_NAME As String = "SalesFactsTable"


'Members functions
Public Function getCities() As Collection
    Set getCities = getBasicMembersFromTable(CITIES_SHEET_NAME, CITIES_TABLE_NAME)
End Function

Public Function getStores() As Collection
    Set getStores = getBasicMembersFromTable(STORES_SHEET_NAME, STORES_TABLE_NAME, True)
End Function

Public Function getProducts() As Collection
    Set getProducts = getBasicMembersFromTable(PRODUCTS_SHEET_NAME, PRODUCTS_TABLE_NAME, True)
End Function

Public Function getProductHierarchy() As HierarchyMember
    Set getProductHierarchy = getHierarchyFromTable(PRODUCTS_SHEET_NAME, PRODUCTS_TABLE_NAME, False, False)
End Function

Public Function getTimeMembers() As Collection
    Set getTimeMembers = getBasicTimeMembersFromTable(TIME_SHEET_NAME, TIME_TABLE_NAME)
End Function




'Facts functions
Public Function getFacts() As Collection
    Set getFacts = getDictsFromFactTable(SALES_FACTS_SHEET_NAME, SALES_FACTS_TABLE_NAME)
End Function

