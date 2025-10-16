Attribute VB_Name = "Specs"
Public Function Specs() As SpecSuite
    Set Specs = New SpecSuite
    Specs.Description = "VBA-XmlConverter"
    
    'On Error Resume Next
    
    Dim XmlString As String
    Dim XmlObject As Dictionary
    'Dim Document As New DOMDocument 'Requires Microsoft XML, v3.0
    Dim Document As New DOMDocument60 'Requires Microsoft XML, v6.0
    
    Document.async = False
    
    ' ============================================= '
    ' ParseXml
    ' ============================================= '
    
    With Specs.It("should parse prolog")
        XmlString = "<?xml version=""1.0""?><!DOCTYPE message [<!ELEMENT message (#PCDATA)>]><message>Howdy!</message>"
        Set XmlObject = XMLConverter.ParseXml(XmlString)

        .Expect(XmlObject("prolog")).ToEqual "<?xml version=""1.0""?>"
    End With

    With Specs.It("should parse doctype")
        XmlString = "<?xml version=""1.0""?><!DOCTYPE message [<!ELEMENT message (#PCDATA)>]><message>Howdy!</message>"
        Set XmlObject = XMLConverter.ParseXml(XmlString)

        Document.LoadXML XmlString

        .Expect(XmlObject("doctype")).ToEqual "<!DOCTYPE message [<!ELEMENT message (#PCDATA)>]>"
    End With
    
    With Specs.It("should parse simple element")
        XmlString = "<messages name=""Tim""><message id=""1"">Howdy!</message><message id=""2"">Howdy 2!</message></messages>"
        Set XmlObject = XMLConverter.ParseXml(XmlString)
        
        Document.LoadXML XmlString
        
        .Expect(Document.nodeName).ToEqual "#document"
        .Expect(Document.documentElement.nodeName).ToEqual "messages"
        .Expect(Document.documentElement.childNodes.Length).ToEqual 2
        .Expect(Document.documentElement.childNodes(0).nodeName).ToEqual "message"
        .Expect(Document.documentElement.childNodes(0).Text).ToEqual "Howdy!"
        .Expect(Document.documentElement.childNodes(0).attributes(0).nodeName).ToEqual "id"
        .Expect(Document.documentElement.childNodes(0).attributes(0).Text).ToEqual "1"
        .Expect(Document.documentElement.childNodes(1).nodeName).ToEqual "message"
        .Expect(Document.documentElement.childNodes(1).Text).ToEqual "Howdy 2!"
        .Expect(Document.documentElement.childNodes(1).attributes(0).nodeName).ToEqual "id"
        .Expect(Document.documentElement.childNodes(1).attributes(0).Text).ToEqual "2"

        .Expect(XmlObject("nodeName")).ToEqual "#document"
        .Expect(XmlObject("childNodes").Count).ToEqual 1
        .Expect(XmlObject("childNodes")(1)("nodeName")).ToEqual "messages"
        .Expect(XmlObject("childNodes")(1)("childNodes").Count).ToEqual 2
        .Expect(XmlObject("childNodes")(1)("childNodes")(1)("nodeName")).ToEqual "message"
        .Expect(XmlObject("childNodes")(1)("childNodes")(1)("text")).ToEqual "Howdy!"
        .Expect(XmlObject("childNodes")(1)("childNodes")(1)("attributes")(1)("name")).ToEqual "id"
        .Expect(XmlObject("childNodes")(1)("childNodes")(1)("attributes")(1)("value")).ToEqual "1"
        .Expect(XmlObject("childNodes")(1)("childNodes")(2)("nodeName")).ToEqual "message"
        .Expect(XmlObject("childNodes")(1)("childNodes")(2)("text")).ToEqual "Howdy 2!"
        .Expect(XmlObject("childNodes")(1)("childNodes")(2)("attributes")(1)("name")).ToEqual "id"
        .Expect(XmlObject("childNodes")(1)("childNodes")(2)("attributes")(1)("value")).ToEqual "2"
    End With
    
    With Specs.It("should parse advanced XML")
        XmlString = "<?xml version=""1.0""?>" & _
            "<ns:Document" & vbNewLine & _
            "    ns:a=""99503""" & vbNewLine & _
            "    ns:b=""1999-10-20""" & vbNewLine & _
            "    xmlns:ns=""http://www.testing.com"">" & vbNewLine & _
            "  <ns:EmptyElement/><ns:EmptyElement ns:c=""123""/><ns:EmptyElement></ns:EmptyElement>" & vbNewLine & _
            "  <ns:Messages>" & vbNewLine & _
            "    <ns:Message ns:d=""2014-11-01"" ns:e=""123"">" & vbNewLine & _
            "      <ns:From><ns:Name>Tim</ns:Name></ns:From>" & vbNewLine & _
            "      <ns:Body>" & vbNewLine & "Howdy!" & vbNewLine & "</ns:Body>" & vbNewLine & _
            "    </ns:Message>" & vbNewLine & _
            "    <ns:Message ns:d=""2014-11-01"" ns:e=""456"">" & vbNewLine & _
            "      <ns:From><ns:Name>Tim</ns:Name></ns:From>" & vbNewLine & _
            "      <ns:Body>" & vbNewLine & "Howdy again!" & vbNewLine & "</ns:Body>" & vbNewLine & _
            "    </ns:Message>" & vbNewLine & _
            "  </ns:Messages>" & vbNewLine & _
            "</ns:Document>"
        
        Set XmlObject = XMLConverter.ParseXml(XmlString)
        
        ' Test document structure
        .Expect(XmlObject("nodeName")).ToEqual "#document"
        .Expect(XmlObject("prolog")).ToEqual "<?xml version=""1.0""?>"
        
        ' Test root element (ns:Document)
        .Expect(XmlObject("childNodes")(1)("nodeName")).ToEqual "ns:Document"
        .Expect(XmlObject("childNodes")(1)("attributes").Count).ToEqual 3
        
        ' Test ns:Document attributes
        .Expect(XmlObject("childNodes")(1)("attributes")(1)("name")).ToEqual "ns:a"
        .Expect(XmlObject("childNodes")(1)("attributes")(1)("value")).ToEqual "99503"
        .Expect(XmlObject("childNodes")(1)("attributes")(2)("name")).ToEqual "ns:b"
        .Expect(XmlObject("childNodes")(1)("attributes")(2)("value")).ToEqual "1999-10-20"
        .Expect(XmlObject("childNodes")(1)("attributes")(3)("name")).ToEqual "xmlns:ns"
        .Expect(XmlObject("childNodes")(1)("attributes")(3)("value")).ToEqual "http://www.testing.com"
        
        ' Test that ns:Document has child elements
        .Expect(XmlObject("childNodes")(1)("childNodes").Count).ToEqual 4
        
        ' Test first empty element
        .Expect(XmlObject("childNodes")(1)("childNodes")(1)("nodeName")).ToEqual "ns:EmptyElement"
        .Expect(XmlObject("childNodes")(1)("childNodes")(1)("text")).ToEqual ""
        
        ' Test ns:Messages element exists
        .Expect(XmlObject("childNodes")(1)("childNodes")(4)("nodeName")).ToEqual "ns:Messages"
        .Expect(XmlObject("childNodes")(1)("childNodes")(4)("childNodes").Count).ToEqual 2
    End With
    
    ' ============================================= '
    ' ConvertToXml
    ' ============================================= '
    
    
    
    ' ============================================= '
    ' Errors
    ' ============================================= '
    
    
    
    InlineRunner.RunSuite Specs
End Function

Public Sub RunSpecs()
    DisplayRunner.IdCol = 1
    DisplayRunner.DescCol = 1
    DisplayRunner.ResultCol = 2
    DisplayRunner.OutputStartRow = 4
    
    DisplayRunner.RunSuite Specs
End Sub

Public Function ToMatchParseError(Actual As Variant, Args As Variant) As Variant
    Dim Partial As String
    Dim Arrow As String
    Dim Message As String
    Dim Description As String
    
    If UBound(Args) < 2 Then
        ToMatchParseError = "Need to pass expected partial, arrow, and message"
    ElseIf Err.Number = 10101 Then
        Partial = Args(0)
        Arrow = Args(1)
        Message = Args(2)
        Description = "Error parsing XML:" & vbNewLine & Partial & vbNewLine & Arrow & vbNewLine & Message
        
        Dim Parts As Variant
        Parts = Split(Err.Description, vbNewLine)
        
        If Parts(1) <> Partial Then
            ToMatchParseError = "Expected " & Parts(1) & " to equal " & Partial
        ElseIf Parts(2) <> Arrow Then
            ToMatchParseError = "Expected " & Parts(2) & " to equal " & Arrow
        ElseIf Parts(3) <> Message Then
            ToMatchParseError = "Expected " & Parts(3) & " to equal " & Message
        ElseIf Err.Description <> Description Then
            ToMatchParseError = "Expected " & Err.Description & " to equal " & Description
        Else
            ToMatchParseError = True
        End If
    Else
        ToMatchParseError = "Expected error number " & Err.Number & " to be 10101"
    End If
End Function
