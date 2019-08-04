Public Class ZefaniaXmlCaption

    Public ReadOnly Property ParentChapter As ZefaniaXmlChapter
    Private CaptionXmlNode As System.Xml.XmlNode

    Friend Sub New(chapter As ZefaniaXmlChapter, xmlNode As System.Xml.XmlNode)
        Me.ParentChapter = chapter
        Me.CaptionXmlNode = xmlNode
    End Sub

    ''' <summary>
    ''' This attribut holds the caption title in long form, e.g. "Bereschit בראשית (Gen 1,1-6,8)"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CaptionTitle As String
        Get
            Return CType(CaptionXmlNode.Attributes("bname").InnerText, String)
        End Get
    End Property

    ''' <summary>
    ''' This attribut holds the caption type, e.g. "x-h2"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property CaptionType As String
        Get
            Return CType(CaptionXmlNode.Attributes("type").InnerText, String)
        End Get
    End Property

    ''' <summary>
    ''' This attribut holds the referenced verse number, e.g. "1"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property VerseNumberReference As Integer
        Get
            Return CType(CaptionXmlNode.Attributes("vref").InnerText, Integer)
        End Get
    End Property

    ''' <summary>
    ''' The referenced verse
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property VerseReference As ZefaniaXmlVerse
        Get
            Return Me.ParentChapter.Verses(Me.VerseNumberReference + 1)
        End Get
    End Property

    ''' <summary>
    ''' This attribut holds the count information, e.g. "146"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Count As Integer
        Get
            Return CType(CaptionXmlNode.Attributes("count").InnerText, Integer)
        End Get
    End Property

    Public Sub ValidateDeeply()
        Try
            'TODO: check caption, e.g. for valid verse reference pointer
        Catch ex As Exception
            Me.ParentChapter.ParentBook.ParentBible.ValidationErrors.Add(ex)
        End Try
    End Sub

End Class
