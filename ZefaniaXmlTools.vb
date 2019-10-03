Option Strict On
Option Explicit On

Module ZefaniaXmlTools

    Private Function EncodeHebrewChars(text As String) As String
        Return text
        'If text = Nothing Then Return text
        'Dim Result As New System.Text.StringBuilder
        'For MyCounter As Integer = 0 To text.Length - 1
        '    'If text(MyCounter) = "'"c OrElse text(MyCounter) = """"c OrElse AscW(text(MyCounter)) > 1450 AndAlso AscW(text(MyCounter)) < 1550 Then
        '    If AscW(text(MyCounter)) > 1450 AndAlso AscW(text(MyCounter)) < 1550 Then
        '        'Hebrew char
        '        Result.Append("&#")
        '        Result.Append(AscW(text(MyCounter)))
        '        Result.Append(";")
        '    Else
        '        'Non Hebrew char
        '        Result.Append(text(MyCounter))
        '    End If
        'Next
        'Return Result.ToString
    End Function

    Private Function EncodeXmlNodeValue(text As String) As String
        If text = Nothing Then Return text
        'return EncodeHebrewChars(text)
        Return text.Replace("&", "&#" & AscW("&"c) & ";").Replace("<", "&#" & AscW("<"c) & ";").Replace(">", "&#" & AscW(">"c) & ";")
    End Function

    Private Function EncodeXmlAttributeValue(text As String) As String
        If text = Nothing Then Return text
        'return EncodeHebrewChars(text)
        Return EncodeXmlNodeValue(text).Replace("'", "&#" & AscW("'"c) & ";").Replace("""", "&#" & AscW(""""c) & ";")
    End Function

    ''' <summary>
    ''' Set an attribute value to an existing XML node
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="name"></param>
    ''' <param name="value"></param>
    Friend Sub WriteAttribute(node As System.Xml.XmlNode, name As String, value As String)
        If value <> Nothing AndAlso value.Contains("'") Then
            Throw New Exception("WARNING RAISED AS EXCEPTION: Char ""'"" (apostroph) not supported in attributes by e.g. SongBeamer; node """ & node.LocalName & """, attribute name """ & name & """, value """ & value & """")
        End If
        If node.Attributes(name) Is Nothing Then
            Dim newAttribute As System.Xml.XmlAttribute = node.OwnerDocument.CreateAttribute(name)
            newAttribute.InnerText = value
            'newAttribute.InnerXml = EncodeXmlAttributeValue(value)
            'newAttribute.Value = value
            node.Attributes.SetNamedItem(newAttribute)
        Else
            node.Attributes(name).InnerText = value
            'node.Attributes(name).InnerXml = EncodeXmlAttributeValue(value)
        End If
    End Sub

    ''' <summary>
    ''' Set a text value to an existing XML node
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="value"></param>
    Friend Sub WriteNodeContent(node As System.Xml.XmlNode, value As String)
        node.InnerText = value
        'node.InnerXml = EncodeXmlNodeValue(value)
    End Sub

End Module
