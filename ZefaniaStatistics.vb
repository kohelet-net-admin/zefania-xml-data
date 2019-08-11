Option Strict On
Option Explicit On

Public Class ZefaniaStatistics

    Public Shared Sub ExportBibleBookNames(exportCsvFilePath As String, bible As ZefaniaXmlBible)
        Dim Result As New DataTable("BibleBooks")
        Result.Columns.Add("Index", GetType(Integer))
        Result.Columns.Add("BookNumber", GetType(Integer))
        Result.Columns.Add("BookName", GetType(String))
        Result.Columns.Add("BookShortName", GetType(String))
        For MyCounter As Integer = 0 To bible.Books.Count - 1
            Dim Row As DataRow = Result.NewRow
            Row(0) = MyCounter
            Row(1) = bible.Books(MyCounter).BookNumber
            Row(2) = bible.Books(MyCounter).BookName
            Row(3) = bible.Books(MyCounter).BookShortName
            Result.Rows.Add(Row)
        Next
        'Result = CompuMaster.Data.DataTables.CreateDataTableClone(Result, "", "BookNumber") 
        CompuMaster.Data.Csv.WriteDataTableToCsvFile(exportCsvFilePath, Result, True, System.Globalization.CultureInfo.InvariantCulture)
    End Sub

    Public Shared Sub ExportBibleStatistics(exportCsvFilePath As String, bible As ZefaniaXmlBible)
        Dim Result As New DataTable("BibleBooks")
        Result.Columns.Add("Index", GetType(Integer))
        Result.Columns.Add("BookNumber", GetType(Integer))
        Result.Columns.Add("BookName", GetType(String))
        Result.Columns.Add("BookShortName", GetType(String))
        Result.Columns.Add("ChaptersCount", GetType(Integer))
        Result.Columns.Add("CaptionsCountTotal", GetType(Integer))
        Result.Columns.Add("CaptionsCountPerChapter", GetType(String))
        Result.Columns.Add("VersesCountTotal", GetType(Integer))
        Result.Columns.Add("VersesCountPerChapter", GetType(String))
        For MyCounter As Integer = 0 To bible.Books.Count - 1
            Dim Row As DataRow = Result.NewRow
            Row(0) = MyCounter
            Row(1) = bible.Books(MyCounter).BookNumber
            Row(2) = bible.Books(MyCounter).BookName
            Row(3) = bible.Books(MyCounter).BookShortName
            Row(4) = bible.Books(MyCounter).Chapters.Count
            Row(5) = CaptionsCountTotalPerChapter(bible.Books(MyCounter))
            Row(6) = CaptionsCountPerChapter(bible.Books(MyCounter))
            Row(7) = VersesCountTotalPerChapter(bible.Books(MyCounter))
            Row(8) = VersesCountPerChapter(bible.Books(MyCounter))
            Result.Rows.Add(Row)
        Next
        'Result = CompuMaster.Data.DataTables.CreateDataTableClone(Result, "", "BookNumber") 
        CompuMaster.Data.Csv.WriteDataTableToCsvFile(exportCsvFilePath, Result, True, System.Globalization.CultureInfo.InvariantCulture)
    End Sub

    Private Shared Function CaptionsCountTotalPerChapter(book As ZefaniaXmlBook) As Integer
        Dim Result As Integer
        For Each Chapter As ZefaniaXmlChapter In book.Chapters
            Result += Chapter.Captions.Count
        Next
        Return Result
    End Function
    Private Shared Function VersesCountTotalPerChapter(book As ZefaniaXmlBook) As Integer
        Dim Result As Integer
        For Each Chapter As ZefaniaXmlChapter In book.Chapters
            Result += Chapter.Verses.Count
        Next
        Return Result
    End Function
    Private Shared Function CaptionsCountPerChapter(book As ZefaniaXmlBook) As String
        Dim Result As New System.Text.StringBuilder
        For Each Chapter As ZefaniaXmlChapter In book.Chapters
            If Chapter.Captions.Count > 0 Then
                If Result.Length > 0 Then Result.Append("|")
                Result.Append(Chapter.ChapterNumber & ":" & Chapter.Captions.Count)
            End If
        Next
        Return Result.ToString
    End Function
    Private Shared Function VersesCountPerChapter(book As ZefaniaXmlBook) As String
        Dim Result As New System.Text.StringBuilder
        For Each Chapter As ZefaniaXmlChapter In book.Chapters
            If Result.Length > 0 Then Result.Append("|")
            Result.Append(Chapter.ChapterNumber & ":" & Chapter.Verses.Count)
        Next
        Return Result.ToString
    End Function

End Class
