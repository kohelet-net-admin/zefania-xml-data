Option Strict On
Option Explicit On

Imports System.Xml
Imports KoheletNetwork

Public Class ZefaniaXmlBibleInfoCache

    Public Sub New(bible As ZefaniaXmlBible)
        'Setup internal fields
        Me.BibleName = bible.BibleName
        Me.BibleRevision = bible.BibleRevision
        Me.BibleStatus = bible.BibleStatus
        Me.BibleType = bible.BibleType
        Me.BibleVersion = bible.BibleVersion
        Me.BibleFreeToEditLicense = bible.BibleFreeToEditLicense
        Me.BibleInfoContributors = bible.BibleInfoContributors
        Me.BibleInfoCoverage = bible.BibleInfoCoverage
        Me.BibleInfoCreator = bible.BibleInfoCreator
        Me.BibleInfoDate = bible.BibleInfoDate
        Me.BibleInfoDescription = bible.BibleInfoDescription
        Me.BibleInfoFormat = bible.BibleInfoFormat
        Me.BibleInfoIdentifier = bible.BibleInfoIdentifier
        Me.BibleInfoLanguage = bible.BibleInfoLanguage
        Me.BibleInfoPublisher = bible.BibleInfoPublisher
        Me.BibleInfoRights = bible.BibleInfoRights
        Me.BibleInfoSources = bible.BibleInfoSource
        Me.BibleInfoSubject = bible.BibleInfoSubject
        Me.BibleInfoTitle = bible.BibleInfoTitle
        Me.BibleInfoType = bible.BibleInfoType
    End Sub

    Public ReadOnly Property BibleName As String
    Public ReadOnly Property BibleRevision As String
    Public ReadOnly Property BibleStatus As String
    Public ReadOnly Property BibleVersion As String
    Public ReadOnly Property BibleType As String

    Public ReadOnly Property BibleFreeToEditLicense As ZefaniaXmlBible.EditLicenseType

    ''' <summary>
    ''' Full bible name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BibleInfoTitle As String

    ''' <summary>
    ''' Unique identifier for bible name
    ''' </summary>
    ''' <returns></returns>
    Public ReadOnly Property BibleInfoIdentifier As String

    Public ReadOnly Property BibleInfoLanguage As String

    Public ReadOnly Property BibleInfoFormat As String

    Public ReadOnly Property BibleInfoDate As String

    Public ReadOnly Property BibleInfoRights As String

    Public ReadOnly Property BibleInfoSources As String

    Public ReadOnly Property BibleInfoSubject As String

    Public ReadOnly Property BibleInfoPublisher As String

    Public ReadOnly Property BibleInfoCreator As String

    Public Property BibleInfoDescription As String

    Public ReadOnly Property BibleInfoType As String

    Public ReadOnly Property BibleInfoCoverage As String

    Public ReadOnly Property BibleInfoContributors As String

End Class
