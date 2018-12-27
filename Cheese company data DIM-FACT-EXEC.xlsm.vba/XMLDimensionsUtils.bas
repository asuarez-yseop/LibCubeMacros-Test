'Doc: Creates a dimension with a plain hierarchy from a collection of basic members, where the parent and root member is the "MEMBER_ALL" entity
Public Sub createBasicHierarchicalDim(dimensionName As String, basicmembers As Collection, Optional isOptionalDim As Boolean = False)
    Dim dimElem As Object
    Dim hierarchiesElem As Object
    Dim rootMember As BasicMember
    Dim members As Collection
    Dim rootElem As Object
    Dim currentElem As Object
    Dim labelElem As Object
    Dim childElem As Object
    Dim valueElem As Object
    
    ' creates <dimensions yid="DIMENSION_SECTOR" />
    createElement element:="dimensions", _
                  parent:=cubeElem, _
                  YID:="DIMENSION_" & UCase(dimensionName), _
                  setElement:=dimElem
                  
                   
    If isOptionalDim Then
        createElement element:="isOptional", _
                  parent:=dimElem, _
                  theText:="True"
    End If
    
    If Right(LCase(dimensionName), 1) = "y" Then
        ' creates <hierarchies yid="HIERARCHY_SECTORS" class="LibCube:Hierarchy">
        createElement element:="hierarchies", _
                      parent:=dimElem, _
                      YID:="HIERARCHY_" & Left(UCase(dimensionName), Len(dimensionName) - 1) & "IES", _
                      yclass:="LibCube:Hierarchy", _
                      setElement:=hierarchiesElem
    Else
        ' creates <hierarchies yid="HIERARCHY_SECTORS" class="LibCube:Hierarchy">
        createElement element:="hierarchies", _
                      parent:=dimElem, _
                      YID:="HIERARCHY_" & UCase(dimensionName) & "S", _
                      yclass:="LibCube:Hierarchy", _
                      setElement:=hierarchiesElem
    End If
    
   
                  
    
    Set rootMember = New BasicMember
    rootMember.id = UCase(dimensionName) & "_ALL"
    
    
    If Right(LCase(dimensionName), 1) = "y" Then
        rootMember.name = "All " & Left(dimensionName, Len(dimensionName) - 1) & "ies"
    Else
        rootMember.name = "All " & dimensionName & "s"
    End If
    
    createElement element:="value", _
                  parent:=hierarchiesElem, _
                  YID:=rootMember.id, _
                  yclass:=dimensionName, _
                  setElement:=rootElem
    
    createElement element:="label", _
              parent:=rootElem, _
              theText:=rootMember.name
    
    For Each BasicMember In basicmembers
        createElement element:="child", _
                  parent:=hierarchiesElem, _
                  yclass:="LibCube:Hierarchy", _
                  setElement:=childElem
        
        createElement element:="value", _
              parent:=childElem, _
              YID:=BasicMember.id, _
              yclass:=dimensionName, _
              setElement:=valueElem
        
        createElement element:="label", _
              parent:=valueElem, _
              theText:=BasicMember.name
              
        For Each AdditionalField In BasicMember.additionalFields
            If AdditionalField.isYClassField = True Then
                createElement element:=AdditionalField.tagName, _
                  parent:=valueElem, _
                  YID:=AdditionalField.value
            Else
                createElement element:=AdditionalField.tagName, _
                  parent:=valueElem, _
                  theText:=AdditionalField.value
            End If
        Next AdditionalField
        
    Next BasicMember
    
End Sub

'Doc: Creates a plain dimension from a collection of basic members
Public Sub createBasicDim(dimensionName As String, basicmembers As Collection, Optional isOptionalDim As Boolean = False)
    Dim dimElem As Object
    Dim hierarchiesElem As Object
    Dim members As Collection
    Dim rootElem As Object
    Dim currentElem As Object
    Dim labelElem As Object
    Dim childElem As Object
    Dim valueElem As Object
    
    ' creates <dimensions yid="DIMENSION_" />
    createElement element:="dimensions", _
                  parent:=cubeElem, _
                  YID:="DIMENSION_" & UCase(dimensionName), _
                  setElement:=dimElem
                  
                   
    If isOptionalDim Then
        createElement element:="isOptional", _
                  parent:=dimElem, _
                  theText:="True"
    End If
    
    
    For Each BasicMember In basicmembers
        
        createElement element:="value", _
              parent:=dimElem, _
              YID:=BasicMember.id, _
              yclass:=dimensionName, _
              setElement:=valueElem
        
        createElement element:="label", _
              parent:=valueElem, _
              theText:=BasicMember.name
              
        For Each AdditionalField In BasicMember.additionalFields
            If AdditionalField.isYClassField = True Then
                createElement element:=AdditionalField.tagName, _
                  parent:=valueElem, _
                  YID:=AdditionalField.value
            Else
                createElement element:=AdditionalField.tagName, _
                  parent:=valueElem, _
                  theText:=AdditionalField.value
            End If
        Next AdditionalField
        
    Next BasicMember
    
End Sub

'Doc: Creates a dimension with a hierarchy using a basic tree structure of hirarchicalEntity members using recursion
Sub createHierarchicalDimension(dimensionTree As HierarchyMember, dimensionName As String, Optional isOptionalDim As Boolean = False)
    Dim dimElem As Object
    Dim hierarchiesElem As Object
    Dim childElem As Object
    Dim valueElem As Object
    Dim rootMember As HierarchyMember
    
    ' creates <dimensions yid="DIMENSION_NAME" />
    createElement element:="dimensions", _
                  parent:=cubeElem, _
                  YID:="DIMENSION_" & toYID(dimensionName), _
                  setElement:=dimElem
                      
    If isOptionalDim Then
        createElement element:="isOptional", _
                  parent:=dimElem, _
                  theText:="True"
    End If
                  
    ' creates <hierarchies yid="HIERARCHY_NAME" class="LibCube:Hierarchy">
    createElement element:="hierarchies", _
                  parent:=dimElem, _
                  YID:="HIERARCHY_" & toYID(dimensionName), _
                  yclass:="LibCube:Hierarchy", _
                  setElement:=hierarchiesElem

    If InStr(dimensionName, "indicator") > 0 Or InStr(dimensionName, "Indicator") > 0 Or InStr(dimensionName, "INDICATOR") > 0 Then
        buildXMLDimensionHierarchy hierarchiesElem, dimensionTree, "LibCube:IndicatorMember"
    Else
        buildXMLDimensionHierarchy hierarchiesElem, dimensionTree, dimensionName
    End If
                  
End Sub

'Doc: Creates a dimension with a hierarchy using a basic tree structure of hirarchicalEntity members using recursion
Sub buildXMLDimensionHierarchy(parentElem As Object, treeNode As HierarchyMember, dimensionName As String)
   Dim currentElem As Object
   Dim labelElem As Object
   Dim childElem As Object
   Dim valueElem As Object
   Dim field As AdditionalField
   Dim currentChild As HierarchyMember
   
   If treeNode.IsRoot Then
       createElement element:="value", _
                 parent:=parentElem, _
                 YID:=treeNode.id, _
                 yclass:=dimensionName, _
                 setElement:=valueElem
       
   Else
     createElement element:="child", _
                 parent:=parentElem, _
                 yclass:="LibCube:Hierarchy", _
                 setElement:=childElem
       
     createElement element:="value", _
             parent:=childElem, _
             YID:=treeNode.id, _
             yclass:=dimensionName, _
             setElement:=valueElem
   End If
   
                 
   createElement element:="label", _
             parent:=valueElem, _
             theText:=treeNode.name, _
             setElement:=labelElem
   
   For Each field In treeNode.additionalFields
       createElement element:=field.tagName, _
             parent:=valueElem, _
             theText:=field.value
   Next field
   
   
   
   For Each child In treeNode.Children
       Set currentChild = child
       If treeNode.IsRoot Then
           buildXMLDimensionHierarchy parentElem, currentChild, dimensionName
       Else
           buildXMLDimensionHierarchy childElem, currentChild, dimensionName
       End If
   Next child
End Sub


Public Sub createBasicTimeDimension(basicTimeMembers As Collection)
    Dim dimElem As Object
    Dim members As Collection
    Dim membersElem As Object
    Dim currentElem As Object
    Dim labelElem As Object
    Dim valueElem As Object
    
    ' creates <dimensions yid="DIMENSION_TIME" />
    createElement element:="dimensions", _
                  parent:=cubeElem, _
                  YID:="DIMENSION_TIME", _
                  setElement:=dimElem
   
    
    For Each member In basicTimeMembers
        createElement element:="members", _
                  parent:=dimElem, _
                  YID:=member.id, _
                  yclass:="LibCube:TimeMember", _
                  setElement:=membersElem
    
        createElement element:="period", _
                  parent:=membersElem, _
                  YID:=member.period
        
        createElement element:="date", _
                  parent:=membersElem, _
                  theText:=Format(member.timeDate, "yyyy/mm/dd")
        
    Next member
   
End Sub


Public Sub createHierarchicalTimeDimension(treeNode As HierarchyTimeMember)
    Dim dimElem As Object
    Dim hierarchiesElem As Object
    
    ' creates <dimensions yid="DIMENSION_TIME" />
    createElement element:="dimensions", _
                  parent:=cubeElem, _
                  YID:="DIMENSION_TIME", _
                  setElement:=dimElem
                  
    ' creates <hierarchies yid="HIERARCHY_TIME" class="LibCube:Hierarchy">
    createElement element:="hierarchies", _
                  parent:=dimElem, _
                  YID:="HIERARCHY_TIME", _
                  yclass:="LibCube:Hierarchy", _
                  setElement:=hierarchiesElem
    
    buildXMLTimeDimensionHierarchy hierarchiesElem, treeNode
    
End Sub

Sub buildXMLTimeDimensionHierarchy(parentElem As Object, treeNode As HierarchyTimeMember)
   Dim currentElem As Object
   Dim labelElem As Object
   Dim childElem As Object
   Dim valueElem As Object
   Dim field As AdditionalField
   Dim currentChild As HierarchyTimeMember
   
   If treeNode.IsRoot Then
       createElement element:="value", _
                 parent:=parentElem, _
                 YID:=treeNode.id, _
                 yclass:="LibCube:TimeMember", _
                 setElement:=valueElem
        
        createElement element:="period", _
            parent:=valueElem, _
            YID:=treeNode.period
        
        createElement element:="date", _
            parent:=valueElem, _
            theText:=Format(treeNode.timeDate, "yyyy/mm/dd")
       
   Else
     createElement element:="child", _
        parent:=parentElem, _
        yclass:="LibCube:Hierarchy", _
        setElement:=childElem
        
        createElement element:="value", _
            parent:=childElem, _
            YID:=treeNode.id, _
            yclass:="LibCube:TimeMember", _
            setElement:=valueElem
        
            createElement element:="period", _
                  parent:=valueElem, _
                  YID:=treeNode.period
                
            createElement element:="date", _
                parent:=valueElem, _
                theText:=Format(treeNode.timeDate, "yyyy/mm/dd")
   End If
  
   For Each child In treeNode.Children
       Set currentChild = child
       If treeNode.IsRoot Then
           buildXMLTimeDimensionHierarchy parentElem, currentChild
       Else
           buildXMLTimeDimensionHierarchy childElem, currentChild
       End If
   Next child
   
End Sub


