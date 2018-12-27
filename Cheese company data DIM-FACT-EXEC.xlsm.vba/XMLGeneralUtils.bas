Public objDom As Object ' DOM object
Public yInstElem As Object
Public cubeElem As Object
Public factElem As Object
Public dimTimeElem As Object
Private caseElem As Object
Private csAttr As Object
Private wAttr As Object
Private yAttr As Object
Private idAttr As Object
Private dataElem As Object

Sub saveXML()
    If Dir(tmpDirectory, vbDirectory) = "" Then ' makes tmp directory if it doesn't exist
       MkDir tmpDirectory
    End If
    objDom.Save (inputXML) ' saves the file
End Sub

Sub createElement(element As String, _
                  parent As Variant, _
                  Optional setElement As Object, _
                  Optional YID As String, _
                  Optional yclass As String, _
                  Optional theText As String, _
                  Optional extraProps As Scripting.Dictionary)
                       
    Set xmlElement = objDom.createElement(element)
    parent.appendChild xmlElement
    
    If Not extraProps Is Nothing Then
        For Each key In extraProps.Keys
            Set xmlAttr = objDom.createAttribute(CStr(key))
            xmlAttr.NodeValue = extraProps(key)
            xmlElement.setAttributeNode xmlAttr
        Next key
    End If
    
    ' adds yid="x"
    If YID <> "" Then
        Set xmlAttr = objDom.createAttribute("yid")
        xmlAttr.NodeValue = YID
        xmlElement.setAttributeNode xmlAttr
    End If

    ' adds yclass="x"
    If yclass <> "" Then
        Set xmlAttr = objDom.createAttribute("yclass")
        xmlAttr.NodeValue = yclass
        xmlElement.setAttributeNode xmlAttr
    End If
   
   ' creates <element>x</element>
    If theText <> "" Then
        xmlElement.text = theText
    End If
    
    'If IsMissing(setElement) Then
        Set setElement = xmlElement
    'End If
    
End Sub

Sub createStaticXML()
    Set objDom = CreateObject("MSXML2.DOMDocument.6.0")

    ' creates <?xml version="1.0" encoding="UTF-8"?>
    Set xmlVersion = objDom.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    objDom.appendChild xmlVersion

    ' creates <y:input xmlns:y="http://www.yseop.com/engine/3" />
    
    Set yInputElem = objDom.createElement("y:input")
    objDom.appendChild yInputElem
    Set yInputAttr = objDom.createAttribute("xmlns:y")
    yInputAttr.NodeValue = "http://www.yseop.com/engine/3"
    yInputElem.setAttributeNode yInputAttr
    Set wAttr = objDom.createAttribute("xmlns:w")
    wAttr.NodeValue = "http://www.yseop.com/widget/1"
    yInputElem.setAttributeNode wAttr

    ' creates <y:datas />
    Set yDatasElem = objDom.createElement("y:data")
    yInputElem.appendChild yDatasElem
    
    ' creates <y:instance yid="theGeneralData" />
    Set yInstElem = objDom.createElement("y:instance")
    yDatasElem.appendChild yInstElem
    Set yInstAttr = objDom.createAttribute("yid")
    yInstAttr.NodeValue = "theGeneralData"
    yInstElem.setAttributeNode yInstAttr
End Sub

Sub createStaticXMLforCase()
    Set objDom = CreateObject("MSXML2.DOMDocument.6.0")
    Set xmlVersion = objDom.createProcessingInstruction("xml", "version='1.0' encoding='UTF-8'")
    objDom.appendChild xmlVersion
    
    Set caseElem = objDom.createElement("cs:case")
    objDom.appendChild caseElem
    
    Set csAttr = objDom.createAttribute("xmlns:cs")
    csAttr.NodeValue = "http://www.yseop.com/case/1"
    caseElem.setAttributeNode csAttr
    
    Set wAttr = objDom.createAttribute("xmlns:w")
    wAttr.NodeValue = "http://www.yseop.com/widget/1"
    caseElem.setAttributeNode wAttr
    
    Set yAttr = objDom.createAttribute("xmlns:y")
    yAttr.NodeValue = "http://www.yseop.com/engine/3"
    caseElem.setAttributeNode yAttr
    
    Set idAttr = objDom.createAttribute("id")
    idAttr.NodeValue = "bruceWayne.xml"
    caseElem.setAttributeNode idAttr
    
    Set caseNameElem = objDom.createElement("cs:name")
    caseNameElem.text = projectName
    caseElem.appendChild caseNameElem
    
    Set dataElem = objDom.createElement("cs:data")
    caseElem.appendChild dataElem
    
    ' creates <y:instance yid="theGeneralData" />
    Set yInstElem = objDom.createElement("y:instance")
    dataElem.appendChild yInstElem
    
    
    Set yInstAttr = objDom.createAttribute("yid")
    yInstAttr.NodeValue = "theGeneralData"
    yInstElem.setAttributeNode yInstAttr
End Sub

Sub setLanguage(theLang As String)
    ' creates <language yid="LANG_xx" />
    createElement element:="language", _
                  parent:=yInstElem, _
                  YID:="LANG_" & theLang
End Sub

Sub createCubeElement()
    ' creates <cube yclass="Cube" />
    createElement element:="cube", _
                  parent:=yInstElem, yclass:="LibCube:Cube", _
                  setElement:=cubeElem
End Sub

Sub createCubeTimeMember(timeMemberID As String, _
                         dateString As String, _
                         periodId As String)

    Dim memberElem As Object

    createElement element:="members", _
                  parent:=dimTimeElem, _
                  YID:=timeMemberID, _
                  yclass:="LibCube:TimeMember", _
                  setElement:=memberElem
                  
    createElement element:="date", _
                  parent:=memberElem, _
                  theText:=dateString
                  
    createElement element:="period", _
                  parent:=memberElem, _
                  YID:=periodId
End Sub

Sub createHierarchyValue(theParent As Object, _
                         ParentId As String, _
                         labelText As String)

    Dim theChild As Object

    createElement element:="value", _
                  parent:=theParent, _
                  YID:=toYID(ParentId), _
                  yclass:="Account", _
                  setElement:=theChild
    createElement element:="label", _
                  parent:=theChild, _
                  theText:=labelText
End Sub


