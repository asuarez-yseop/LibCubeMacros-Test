
Public dataSheet As Worksheet ' data worksheet

Public Const tmpDirectory = "C:\tmp\" ' temp directory for input

Public Const inputXML = tmpDirectory & "input.xml" ' input XML file

Public Const outputHTML = tmpDirectory & "output.html" ' output HTML file

Public Const outputXML = tmpDirectory & "output.xml" ' output XML file

Public Const pptTemplate = tmpDirectory & "template.pptx" ' PowerPoint template

Public Const wordTemplate = tmpDirectory & "templateBotW.docx"  ' Word template

Public Const projectName = "LibCube Basic Example" ' the project name


' login/password
'Public Const login = "trianz"
'Public Const password = "trianz2018"

' base URL for POST method
'Public Const baseURL = "https://pejaraba.studio.yseop-hosting.com"
'Public Const baseURL = "https://yet-us.yseop-hosting.com"
'Public Const baseURL = "http://localhost:8080"
'Public Const transform = "" '"htmlWithData"

' generic error message
Public Const genericError = "Text generation failed. " & _
                            "Please ensure your data is correct, " & _
                            "you are connected to the Internet and your " & _
                            "license code has been inputted correctly." & vbCrLf & _
                            vbCrLf



