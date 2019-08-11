Option Strict On
Option Explicit On

Public Class ZefaniaXmlBook

    Public ReadOnly Property ParentBible As ZefaniaXmlBible
    Friend BookXmlNode As System.Xml.XmlNode

    Friend Sub New(bible As ZefaniaXmlBible, xmlNode As System.Xml.XmlNode)
        Me.ParentBible = bible
        Me.BookXmlNode = xmlNode
    End Sub

    ''' <summary>
    ''' Move this book to a new position in the bible collection
    ''' </summary>
    ''' <param name="insertBefore"></param>
    Public Sub MoveBookPositionInBible(insertBefore As ZefaniaXmlBook)
        'reorder book
        If insertBefore IsNot Nothing Then
            'insert before given node
            BookXmlNode.ParentNode.InsertBefore(Me.BookXmlNode, insertBefore.BookXmlNode)
        Else
            'insert after all nodes
            BookXmlNode.ParentNode.InsertAfter(Me.BookXmlNode, BookXmlNode.ParentNode.LastChild)
        End If
        'reset books collection cache of ParentBible
        ParentBible.ResetBooksCache()
    End Sub

    ''' <summary>
    ''' This attribut should hold the book name in long form, e.g. "Genesis"
    ''' </summary>
    ''' <returns></returns>
    Public Property BookName As String
        Get
            If BookXmlNode.Attributes("bname") Is Nothing Then
                Throw New NullReferenceException("BookName attribute ""bname"" not found")
            Else
                Return CType(BookXmlNode.Attributes("bname").InnerText, String)
            End If
        End Get
        Set(value As String)
            WriteAttribute(BookXmlNode, "bname", value)
        End Set
    End Property

    ''' <summary>
    ''' This attribute holds the book book name in short form, e.g. "Gen"
    ''' </summary>
    ''' <returns></returns>
    Public Property BookShortName As String
        Get
            Return CType(BookXmlNode.Attributes("bsname").InnerText, String)
        End Get
        Set(value As String)
            WriteAttribute(BookXmlNode, "bsname", value)
        End Set
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
            For MyCounter As Integer = 0 To Me.Chapters.Count - 1
                Me.Chapters(MyCounter).ValidateDeeply()
            Next
        Catch ex As Exception
            Me.ParentBible.ValidationErrors.Add(ex)
        End Try
    End Sub

    ''' <summary>
    ''' List of all available chapters
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Chapters As List(Of ZefaniaXmlChapter)
        Get
            Static Result As List(Of ZefaniaXmlChapter)
            If Result Is Nothing Then
                Result = New List(Of ZefaniaXmlChapter)
                Try
                    'Me.ValidateXmlData()
                    Dim AllChapters As System.Xml.XmlNodeList = Me.BookXmlNode.SelectNodes("CHAPTER")
                    If AllChapters.Count = 0 Then
                        Me.ParentBible.ValidationErrors.Add(New Exception("Chapter collection for book " & Me.BookName & " is empty"))
                    Else
                        For Each ChapterNode As System.Xml.XmlNode In AllChapters
                            Result.Add(New ZefaniaXmlChapter(Me, ChapterNode))
                        Next
                    End If
                Catch ex As Exception
                    Me.ParentBible.ValidationErrors.Add(New Exception("Chapter collection for book " & Me.BookName & " not accessible", ex))
                End Try
            End If
            Return Result
        End Get
    End Property

End Class
