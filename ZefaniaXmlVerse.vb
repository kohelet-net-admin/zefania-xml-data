Option Strict On
Option Explicit On

Public Class ZefaniaXmlVerse

    Public ReadOnly Property ParentChapter As ZefaniaXmlChapter
    Private VerseXmlNode As System.Xml.XmlNode

    Friend Sub New(chapter As ZefaniaXmlChapter, xmlNode As System.Xml.XmlNode)
        Me.ParentChapter = chapter
        Me.VerseXmlNode = xmlNode
    End Sub

    ''' <summary>
    ''' This verse content as plain text
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ContentText As String
        Get
            Return CType(VerseXmlNode.InnerText, String)
        End Get
    End Property

    ''' <summary>
    ''' This verse content with Zefania XML markup
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ContentXml As String
        Get
            Return CType(VerseXmlNode.InnerXml, String)
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
