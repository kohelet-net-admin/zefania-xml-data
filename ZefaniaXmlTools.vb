Option Strict On
Option Explicit On

Module ZefaniaXmlTools

    ''' <summary>
    ''' Set an attribute value to an existing XML node
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="name"></param>
    ''' <param name="value"></param>
    Friend Sub WriteAttribute(node As System.Xml.XmlNode, name As String, value As String)
        If node.Attributes(name) Is Nothing Then
            Dim newAttribute As System.Xml.XmlAttribute = node.OwnerDocument.CreateAttribute(name)
            newAttribute.Value = value
            node.Attributes.SetNamedItem(newAttribute)
        Else
            node.Attributes(name).InnerText = value
        End If
    End Sub

    ''' <summary>
    ''' Set a text value to an existing XML node
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="value"></param>
    Friend Sub WriteNodeContent(node As System.Xml.XmlNode, value As String)
        node.InnerText = value
    End Sub

End Module
