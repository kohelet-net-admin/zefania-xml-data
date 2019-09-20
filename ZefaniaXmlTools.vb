Option Strict On
Option Explicit On

Module ZefaniaXmlTools

    Friend Function EncodeHebrewChars(text As String) As String
        Return text
        If text = Nothing Then Return text
        Dim Result As New System.Text.StringBuilder
        For MyCounter As Integer = 0 To text.Length - 1
            If AscW(text(MyCounter)) > 1450 AndAlso AscW(text(MyCounter)) < 1550 Then
                'Hebrew char
                Result.Append("&#")
                Result.Append(AscW(text(MyCounter)))
                Result.Append(";")
            Else
                'Non Hebrew char
                Result.Append(text(MyCounter))
            End If
        Next
        Return Result.ToString
    End Function

    ''' <summary>
    ''' Set an attribute value to an existing XML node
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="name"></param>
    ''' <param name="value"></param>
    Friend Sub WriteAttribute(node As System.Xml.XmlNode, name As String, value As String)
        If node.Attributes(name) Is Nothing Then
            Dim newAttribute As System.Xml.XmlAttribute = node.OwnerDocument.CreateAttribute(name)
            newAttribute.InnerText = EncodeHebrewChars(value)
            'newAttribute.Value = value
            node.Attributes.SetNamedItem(newAttribute)
        Else
            node.Attributes(name).InnerText = EncodeHebrewChars(value)
        End If
    End Sub

    ''' <summary>
    ''' Set a text value to an existing XML node
    ''' </summary>
    ''' <param name="node"></param>
    ''' <param name="value"></param>
    Friend Sub WriteNodeContent(node As System.Xml.XmlNode, value As String)
        node.InnerText = EncodeHebrewChars(value)
    End Sub

End Module
