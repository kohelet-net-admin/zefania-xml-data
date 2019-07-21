Public Class ZefaniaXmlBook

    Private ParentBible As ZefaniaXmlBible
    Private BookXmlNode As System.Xml.XmlNode

    Friend Sub New(bible As ZefaniaXmlBible, xmlNode As System.Xml.XmlNode)
        Me.ParentBible = bible
        Me.BookXmlNode = xmlNode
    End Sub

    Public ReadOnly Property BookName As String
        Get
            Try
                'Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("title").Item(0).InnerText
                Return CType(BookXmlNode.Attributes("bname").InnerText, String)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BookShortName As String
        Get
            Try
                'Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("title").Item(0).InnerText
                Return CType(BookXmlNode.Attributes("bsname").InnerText, String)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BookNumber As Integer
        Get
            Try
                'Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("title").Item(0).InnerText
                Return CType(BookXmlNode.Attributes("bnumber").InnerText, Integer)
            Catch
                Return Nothing
            End Try
        End Get
    End Property

End Class
