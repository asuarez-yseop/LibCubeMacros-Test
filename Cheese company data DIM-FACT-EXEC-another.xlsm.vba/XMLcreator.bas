
Sub createXMLforPost()
    Application.ScreenUpdating = False
    Set currentSheet = ActiveWorkbook.ActiveSheet

    createStaticXML
    setLanguage ("en")
    createCubeElement
    createDimensions
    createFacts
    'createExtraFields
    saveXML
    
    currentSheet.Select
    Application.ScreenUpdating = True
End Sub

Sub createXMLforCase()
    Application.ScreenUpdating = False
    Set currentSheet = ActiveWorkbook.ActiveSheet
     
    createStaticXMLforCase
    setLanguage ("en")
    createCubeElement
    createDimensions
    createFacts
    'createExtraFields
    saveXML
    
    currentSheet.Select
    Application.ScreenUpdating = True
End Sub

