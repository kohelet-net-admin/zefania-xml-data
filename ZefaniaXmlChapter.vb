Option Strict On
Option Explicit On

Imports KoheletNetwork

Public Class ZefaniaXmlChapter

    Public ReadOnly Property ParentBook As ZefaniaXmlBook
    Private ChapterXmlNode As System.Xml.XmlNode

    Friend Sub New(book As ZefaniaXmlBook, xmlNode As System.Xml.XmlNode)
        Me.ParentBook = book
        Me.ChapterXmlNode = xmlNode
    End Sub

    ''' <summary>
    ''' This attribut should hold the chapter number, e.g. "1"
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ChapterNumber As String
        Get
            Return Integer.Parse(ChapterXmlNode.Attributes("cnumber").InnerText).ToString
        End Get
    End Property

    ''' <summary>
    ''' List of all available captions
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Captions As List(Of ZefaniaXmlCaption)
        Get
            Static Result As List(Of ZefaniaXmlCaption)
            If Result Is Nothing Then
                Result = New List(Of ZefaniaXmlCaption)
                Try
                    'Me.ValidateXmlData()
                    Dim AllCaptions As System.Xml.XmlNodeList = Me.ChapterXmlNode.SelectNodes("CAPTION")
                    If AllCaptions.Count = 0 Then
                        Me.ParentBook.ParentBible.ValidationErrors.Add(New Exception("Caption collection for book " & Me.ParentBook.BookName & ", chapter " & Me.ChapterNumber & " is empty"))
                    Else
                        For Each CaptionNode As System.Xml.XmlNode In AllCaptions
                            Result.Add(New ZefaniaXmlCaption(Me, CaptionNode))
                        Next
                    End If
                Catch ex As Exception
                    Me.ParentBook.ParentBible.ValidationErrors.Add(New Exception("Caption collection for book " & Me.ParentBook.BookName & ", chapter " & Me.ChapterNumber & " not accessible", ex))
                End Try
            End If
            Return Result
        End Get
    End Property

    ''' <summary>
    ''' List of all available verses
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Verses As List(Of ZefaniaXmlVerse)
        Get
            Static Result As List(Of ZefaniaXmlVerse)
            If Result Is Nothing Then
                Result = New List(Of ZefaniaXmlVerse)
                Try
                    'Me.ValidateXmlData()
                    Dim AllVerses As System.Xml.XmlNodeList = Me.ChapterXmlNode.SelectNodes("VERS")
                    If AllVerses.Count = 0 Then
                        Me.ParentBook.ParentBible.ValidationErrors.Add(New Exception("Verse collection for book " & Me.ParentBook.BookName & ", chapter " & Me.ChapterNumber & " is empty"))
                    Else
                        For Each VerseNode As System.Xml.XmlNode In AllVerses
                            Result.Add(New ZefaniaXmlVerse(Me, VerseNode))
                        Next
                    End If
                Catch ex As Exception
                    Me.ParentBook.ParentBible.ValidationErrors.Add(New Exception("Verse collection for book " & Me.ParentBook.BookName & ", chapter " & Me.ChapterNumber & " not accessible", ex))
                End Try
            End If
            Return Result
        End Get
    End Property

    Public Sub ValidateDeeply()
        Try
            'TODO: interate through captions collection
            For MyCounter As Integer = 0 To Me.Captions.Count - 1
                Me.Captions(MyCounter).ValidateDeeply()
            Next
            'TODO: interate through verses collection
            For MyCounter As Integer = 0 To Me.Verses.Count - 1
                Me.Verses(MyCounter).ValidateDeeply()
            Next
        Catch ex As Exception
            Me.ParentBook.ParentBible.ValidationErrors.Add(ex)
        End Try
    End Sub

End Class
