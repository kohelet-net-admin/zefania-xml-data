Public Class ZefaniaXmlBook

    Private ParentBible As ZefaniaXmlBible
    Private BookXmlNode As System.Xml.XmlNode

    Friend Sub New(bible As ZefaniaXmlBible, xmlNode As System.Xml.XmlNode)
        Me.ParentBible = bible
        Me.BookXmlNode = xmlNode
    End Sub

    ''' <summary>
    ''' Move this book to a new position in the bible collection
    ''' </summary>
    ''' <param name="insertBefore"></param>
    Public Sub MoveBookPositionInBible(insertBefore As ZefaniaXmlBook)
        BookXmlNode.ParentNode.InsertBefore(Me.BookXmlNode, insertBefore.BookXmlNode)
    End Sub

    ''' <summary>
    ''' This attribut should hold the book name in long form e.x. "Genesis"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BookName As String
        Get
            Return CType(BookXmlNode.Attributes("bname").InnerText, String)
        End Get
    End Property

    ''' <summary>
    ''' This attribute holds the book book name in short form e.x. "Gen"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BookShortName As String
        Get
            Return CType(BookXmlNode.Attributes("bsname").InnerText, String)
        End Get
    End Property

    ''' <summary>
    ''' A number which is unambiguous for a certain bible book
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BookNumber As Integer
        Get
            Return CType(BookXmlNode.Attributes("bnumber").InnerText, Integer)
        End Get
    End Property

    Public Sub ValidateDeeply()
        Try
            'TODO: interate through chapters collection
        Catch ex As Exception
            Me.ParentBible.ValidationErrors.Add(ex)
        End Try
    End Sub

End Class
