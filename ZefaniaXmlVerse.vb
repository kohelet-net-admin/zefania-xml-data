Public Class ZefaniaXmlVerse

    Public ReadOnly Property ParentChapter As ZefaniaXmlChapter
    Private VerseXmlNode As System.Xml.XmlNode

    Friend Sub New(chapter As ZefaniaXmlChapter, xmlNode As System.Xml.XmlNode)
        Me.ParentChapter = chapter
        Me.VerseXmlNode = xmlNode
    End Sub

    ''' <summary>
    ''' This attribut holds the caption title in long form, e.g. "Bereschit בראשית (Gen 1,1-6,8)"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Content As String
        Get
            Return CType(VerseXmlNode.Attributes("").InnerText, String)
        End Get
    End Property

    ''' <summary>
    ''' This attribut holds the referenced verse number, e.g. "1"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property VerseNumber As Integer
        Get
            Return CType(VerseXmlNode.Attributes("vnumber").InnerText, Integer)
        End Get
    End Property

    Public Sub ValidateDeeply()
        Try
            'TODO: check verse data, e.g. for valid verse content
        Catch ex As Exception
            Me.ParentChapter.ParentBook.ParentBible.ValidationErrors.Add(ex)
        End Try
    End Sub

End Class
