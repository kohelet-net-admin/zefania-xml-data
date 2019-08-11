﻿Option Strict On
Option Explicit On

Imports System
Imports System.Xml
Imports System.Xml.Schema
Imports System.Xml.XPath

Public Class ZefaniaXmlValidation

    Public Shared Function MD5FileHash(ByVal xmlFile As String) As String
        Dim MD5 As New System.Security.Cryptography.MD5CryptoServiceProvider
        Dim Hash As Byte()
        Dim Result As String = ""
        Dim Tmp As String = ""

        Dim FN As New System.IO.FileStream(xmlFile, System.IO.FileMode.Open, System.IO.FileAccess.Read, System.IO.FileShare.Read, 8192)
        MD5.ComputeHash(FN)
        FN.Close()

        Hash = MD5.Hash
        For i As Integer = 0 To Hash.Length - 1
            Tmp = Hex(Hash(i))
            If Len(Tmp) = 1 Then Tmp = "0" & Tmp
            Result += Tmp
        Next
        Return Result
    End Function

    'Public Enum ValidationResult As Byte
    '    NoWarnings = 0
    '    HasWarnings = 1
    '    HasErrors = 2
    'End Enum

    'Public Shared Function ValidateXml(targetNamespace As String, schemaFile As String, xmlFile As String) As ValidationResult
    '    Dim reader As XmlReader = Nothing
    '    Try
    '        Dim settings As XmlReaderSettings = New XmlReaderSettings()
    '        settings.Schemas.Add(targetNamespace, schemaFile)
    '        settings.ValidationType = ValidationType.Schema

    '        reader = XmlReader.Create(xmlFile, settings)
    '        Dim document As XmlDocument = New XmlDocument()
    '        document.Load(reader)

    '        Dim eventHandler As ValidationEventHandler = New ValidationEventHandler(AddressOf ValidationEventHandler)
    '        document.Validate(eventHandler)
    '    Catch ex As Exception
    '        ValidationErrors.Add(ex)
    '        Return ValidationResult.HasErrors
    '    Finally
    '        If Not reader Is Nothing Then
    '            reader.Close()
    '            reader.Dispose()
    '        End If
    '    End Try
    '    If ValidationErrors.Count > 0 Then
    '        Return ValidationResult.HasErrors
    '    ElseIf ValidationWarnings.Count > 0 Then
    '        Return ValidationResult.HasWarnings
    '    Else
    '        Return ValidationResult.NoWarnings
    '    End If
    'End Function

    'Public Shared ReadOnly Property ValidationWarnings As New List(Of Exception)
    'Public Shared ReadOnly Property ValidationErrors As New List(Of Exception)

    'Private Shared Sub ValidationEventHandler(ByVal sender As Object, ByVal e As ValidationEventArgs)
    '    Select Case e.Severity
    '        Case XmlSeverityType.Error
    '            ValidationErrors.Add(e.Exception)
    '        Case XmlSeverityType.Warning
    '            ValidationWarnings.Add(e.Exception)
    '    End Select
    'End Sub

End Class