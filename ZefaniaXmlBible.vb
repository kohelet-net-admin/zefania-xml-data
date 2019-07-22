Option Strict On
Option Explicit On

Imports System.Xml

Public Class ZefaniaXmlBible

    ''' <summary>
    ''' The directory path where the XSD schema files for ZefaniaXml are stored
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property XsdDirectory As String
    Private XmlFilePath As String
    Private ReadOnly Property XmlDocument As XmlDocument

    Public Sub New(filePath As String, xsdDirectory As String)
        'Setup internal fields
        Me.XmlFilePath = filePath
        Me.XsdDirectory = xsdDirectory
        'Load XmlDocument
        Dim reader As XmlReader = Nothing
        reader = XmlReader.Create(filePath)
        Dim document As XmlDocument = New XmlDocument()
        document.Load(reader)
        XmlDocument = document
        reader.Close()
        reader.Dispose()
    End Sub


    Public ReadOnly Property SchemaName As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).Attributes("xsi:noNamespaceSchemaLocation").Value
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleName As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).Attributes("biblename").Value
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleRevision As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).Attributes("revision").Value
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleStatus As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).Attributes("status").Value
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleVersion As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).Attributes("version").Value
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleType As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).Attributes("type").Value
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleWithoutCopyright As Boolean
        Get
            If Me.BibleInfoRights <> Nothing AndAlso (Me.BibleInfoRights.ToLowerInvariant = "god" OrElse Me.BibleInfoRights.ToLowerInvariant.Contains("public domain")) Then
                Return True
            ElseIf Me.BibleInfoTitle <> Nothing Then
                'If Me.BibleInfoTitle Then
                Dim rgx As System.Text.RegularExpressions.Regex
                rgx = New System.Text.RegularExpressions.Regex("^.*\d{4}.*(\d{4}).*$")
                If rgx.IsMatch(Me.BibleInfoTitle) Then
                    Return BibleWithoutCopyRightRegExCheck(rgx)
                Else
                    rgx = New System.Text.RegularExpressions.Regex("^.*(\d{4}).*$")
                    If rgx.IsMatch(Me.BibleInfoTitle) Then
                        Return BibleWithoutCopyRightRegExCheck(rgx)
                    Else
                        Return False
                    End If
                End If
            Else
                Return False
            End If
        End Get
    End Property

    Private Function BibleWithoutCopyRightRegExCheck(rgx As System.Text.RegularExpressions.Regex) As Boolean
        If rgx.IsMatch(Me.BibleInfoTitle) Then
            Dim Year As String = rgx.Match(Me.BibleInfoTitle).Groups(1).Value
            If Integer.Parse(Year) = 0 Then
                Return False
            ElseIf Integer.Parse(Year) < Now.Year - 80 - 1 Then
                Return True
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function
    ''' <summary>
    ''' Full bible name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BibleInfoTitle As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("title").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Unique identifier for bible name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BibleInfoIdentifier As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("identifier").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleInfoLanguage As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("language").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    Public ReadOnly Property BibleInfoFormat As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("format").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoDate As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("date").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoRights As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("rights").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoSources As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("sources").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoSubject As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("subject").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoPublisher As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("publisher").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoCreator As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("creator").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoDescription As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("description").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoType As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("type").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property
    Public ReadOnly Property BibleInfoCoverage As String
        Get
            Try
                Return XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("INFORMATION").Item(0).SelectNodes("coverage").Item(0).InnerText
            Catch
                Return Nothing
            End Try
        End Get
    End Property

    ''' <summary>
    ''' Check for the XML schema name and validate accordingly
    ''' </summary>
    Private Sub ValidateXmlData()
        'Reset ValidationXml results
        ValidationXmlErrors.Clear()
        ValidationXmlWarnings.Clear()
        'Consider schema and validate document
        If Me.SchemaName = "zef2005.xsd" Then
            'Me.ValidateXmlDataWithSchema("http://www.bgfdb.de/zefaniaxml/2014/", "zef2014.xsd")
            Me.ValidateXmlDataWithSchema("http://www.bgfdb.de/zefaniaxml/2005/", Me.SchemaName)
        ElseIf Me.SchemaName = "zef2014.xsd" Then
            Me.ValidateXmlDataWithSchema("http://www.bgfdb.de/zefaniaxml/2014/", Me.SchemaName)
        ElseIf Me.SchemaName = "" Then 'empty schema
            Throw New ArgumentNullException("xsi:noNamespaceSchemaLocation", "Schema declaration required")
        Else 'unknown schema
            Throw New NotSupportedException("Schema not supported: " & Me.SchemaName)
        End If
    End Sub

    ''' <summary>
    ''' Check the XmlDocument based on its named schema
    ''' </summary>
    Private Sub ValidateXmlDataWithSchema(targetNamespace As String, schemaName As String)
        Dim ZefaniaXmlSchema As System.Xml.Schema.XmlSchema
        ZefaniaXmlSchema = System.Xml.Schema.XmlSchema.Read(System.Xml.XmlReader.Create(System.IO.Path.Combine(Me.XsdDirectory, schemaName)), Nothing)
        ZefaniaXmlSchema.TargetNamespace = targetNamespace
        If Me.XmlDocument.Schemas.Contains(targetNamespace) = False Then
            Me.XmlDocument.Schemas.Add(ZefaniaXmlSchema)
        End If
        Dim eventHandler As System.Xml.Schema.ValidationEventHandler = New System.Xml.Schema.ValidationEventHandler(AddressOf ValidationEventHandler)
        Me.XmlDocument.Validate(eventHandler)
        If ValidationXmlErrors.Count > 0 Then
            Dim sb As New System.Text.StringBuilder()
            sb.AppendLine()
            sb.AppendLine(">>>")
            For MyCounter As Integer = 0 To ValidationXmlErrors.Count - 1
                sb.AppendLine(ValidationXmlErrors(MyCounter).ToString)
            Next
            sb.AppendLine("<<<")
            Throw New Exception("XmlValidation failed with errors:" & sb.ToString)
        End If
    End Sub

    'Public Function ValidateXml() As ZefaniaXmlValidation.ValidationResult
    '    Return ZefaniaXmlValidation.ValidateXml("http://www.bgfdb.de/zefaniaxml/2014/", System.IO.Path.Combine(Me.XsdDirectory, Me.SchemaName), Me.XmlFilePath)
    'End Function

    Public Enum ValidationResult As Byte
        NoWarnings = 0
        HasWarnings = 1
        HasErrors = 2
    End Enum

    ''' <summary>
    ''' Is the XmlDocument successfully validated against the XmlSchema
    ''' </summary>
    ''' <remarks>, Success on warnings or without any</remarks>
    ''' <returns>False on exceptions or validation errors, True on warnings, True without any errors or warnings </returns>
    Public Function IsValidXml() As ValidationResult
        Try
            Me.ValidateXmlData()
            If ValidationXmlErrors.Count > 0 Then
                Return ValidationResult.HasErrors
            ElseIf ValidationXmlWarnings.Count > 0 Then
                Return ValidationResult.HasWarnings
            Else
                Return ValidationResult.NoWarnings
            End If
        Catch ex As Exception
            ValidationXmlErrors.Insert(0, ex)
            Return ValidationResult.HasErrors
        End Try
    End Function

    ''' <summary>
    ''' List of all available books
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property Books As List(Of ZefaniaXmlBook)
        Get
            Static Result As List(Of ZefaniaXmlBook)
            If Result Is Nothing Then
                Result = New List(Of ZefaniaXmlBook)
                Try
                    'Me.ValidateXmlData()
                    Dim AllBooks As XmlNodeList = XmlDocument.SelectNodes("/XMLBIBLE").Item(0).SelectNodes("BIBLEBOOK")
                    If AllBooks.Count = 0 Then
                        Me.ValidationErrors.Add(New Exception("Books collection is empty"))
                    Else
                        For Each BookNode As XmlNode In AllBooks
                            Result.Add(New ZefaniaXmlBook(Me, BookNode))
                        Next
                    End If
                Catch ex As Exception
                    Me.ValidationErrors.Add(New Exception("Books collection not accessible", ex))
                End Try
            End If
            Return Result
        End Get
    End Property

    ''' <summary>
    ''' Validate the XmlDocument as well as internal data structure inclusive all child objects
    ''' </summary>
    Public Sub ValidateDeeply()
        'Reset errors list
        Me.ValidationErrors.Clear()
        'Validate schema
        Try
            Me.ValidateXmlData()
        Catch ex As Exception
            Me.ValidationErrors.Add(ex)
        End Try
        'Check for (duplicate) book no. assignment
        With True
            Dim BookNos As New List(Of Integer)
            For MyCounter As Integer = 0 To Me.Books.Count - 1
                Try
                    Dim BookNo As Integer = Me.Books(MyCounter).BookNumber
                    If BookNo <= 0 Then
                        Me.ValidationErrors.Add(New Exception("Invalid book number for book index " & MyCounter))
                    ElseIf BookNos.Contains(BookNo) Then
                        Me.ValidationErrors.Add(New Exception("Book number " & BookNo & " identified as duplicate for book index " & MyCounter))
                    Else
                        BookNos.Add(BookNo)
                    End If
                Catch ex As Exception
                    Me.ValidationErrors.Add(New Exception("BookNumber not available for book index " & MyCounter, ex))
                End Try
            Next
        End With
        'Check for (duplicate) book name assignment
        With True
            Dim BookNames As New List(Of String)
            For MyCounter As Integer = 0 To Me.Books.Count - 1
                Try
                    Dim BookName As String = Me.Books(MyCounter).BookName
                    If Trim(BookName) = Nothing Then
                        Me.ValidationErrors.Add(New Exception("Missing book name for book index " & MyCounter))
                    ElseIf BookNames.Contains(BookName) Then
                        Me.ValidationErrors.Add(New Exception("Book name " & BookName & " identified as duplicate for book index " & MyCounter))
                    ElseIf BookName.Contains("?") Then
                        Me.ValidationErrors.Add(New Exception("Invalid characters in book name " & BookName & " for book index " & MyCounter))
                    Else
                        BookNames.Add(BookName)
                    End If
                Catch ex As Exception
                    Me.ValidationErrors.Add(New Exception("Book name not available for book index " & MyCounter, ex))
                End Try
            Next
        End With
        'Check for (duplicate) book shortname assignment
        With True
            Dim BookShortNames As New List(Of String)
            For MyCounter As Integer = 0 To Me.Books.Count - 1
                Try
                    Dim BookShortName As String = Me.Books(MyCounter).BookShortName
                    If Trim(BookShortName) = Nothing Then
                        Me.ValidationErrors.Add(New Exception("Missing book short name for book index " & MyCounter))
                    ElseIf BookShortNames.Contains(BookShortName) Then
                        Me.ValidationErrors.Add(New Exception("Book short name " & BookShortName & " identified as duplicate for book index " & MyCounter))
                    ElseIf BookShortName.Contains("?") Then
                        Me.ValidationErrors.Add(New Exception("Invalid characters in book short name " & BookShortName & " for book index " & MyCounter))
                    Else
                        BookShortNames.Add(BookShortName)
                    End If
                Catch ex As Exception
                    Me.ValidationErrors.Add(New Exception("Book short name not available for book index " & MyCounter, ex))
                End Try
            Next
        End With
    End Sub

    ''' <summary>
    ''' Validation errors (not warnings) detected by method ValidateDeeply()
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property ValidationErrors As New List(Of Exception)
    ''' <summary>
    ''' Validation warnings detected by method IsValidXml (only XmlDocument schema validation)
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property ValidationXmlWarnings As New List(Of Exception)
    ''' <summary>
    ''' Validation errors detected by method IsValidXml (only XmlDocument schema validation)
    ''' </summary>
    ''' <returns></returns>
    Public Shared ReadOnly Property ValidationXmlErrors As New List(Of Exception)

    ''' <summary>
    ''' Catch validation events for warnings or errors
    ''' </summary>
    ''' <param name="sender"></param>
    ''' <param name="e"></param>
    Private Shared Sub ValidationEventHandler(ByVal sender As Object, ByVal e As System.Xml.Schema.ValidationEventArgs)
        Select Case e.Severity
            Case System.Xml.Schema.XmlSeverityType.Error
                If ValidationXmlErrors.ConvertAll(Of String)(ValidationEventHandlerExceptionToString()).Contains(e.ToString) = False Then
                    ValidationXmlErrors.Add(e.Exception)
                End If
            Case System.Xml.Schema.XmlSeverityType.Warning
                If ValidationXmlWarnings.ConvertAll(Of String)(ValidationEventHandlerExceptionToString()).Contains(e.ToString) = False Then
                    ValidationXmlWarnings.Add(e.Exception)
                End If
        End Select

    End Sub

    ''' <summary>
    ''' Convert an exception into a string
    ''' </summary>
    ''' <returns></returns>
    Private Shared Function ValidationEventHandlerExceptionToString() As Converter(Of Exception, String)
        Return New Converter(Of Exception, String)(Function(ex As Exception)
                                                       Return ex.ToString()
                                                   End Function)
    End Function

End Class
