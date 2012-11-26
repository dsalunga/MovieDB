Imports System.Data.OleDb
Imports Microsoft.VisualBasic

Module ModuleShared
    Public currentRecord As Long
    Public ARecordSelected As Boolean = False
    Public Const conStr As String = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=Database.mdb;Mode=Share Deny None"
    Public AddNewRecord As Boolean = False
    Public lastBorrower As String

    Public Function FormatDate(ByVal dateToFormat As Date) As String
        Dim dateStr As String

        dateStr = MonthName(Month(dateToFormat)) & " " & Day(dateToFormat).ToString & ", " & Year(dateToFormat).ToString

        FormatDate = dateStr
    End Function

    Public Sub ExecuteCommand(ByVal myExecuteQuery As String)
        Dim myConnection As New OleDbConnection(conStr)
        Dim myCommand As New OleDbCommand(myExecuteQuery, myConnection)
        myCommand.Connection.Open()
        myCommand.ExecuteNonQuery()

        myConnection.Close()
    End Sub

    Public Enum RecordColumns
        Title = 0
        Starring = 1
        Director = 2
        Producer = 3
        Category = 4
        Rating = 5
        Year = 6
        RecordDate = 7
        Type = 8
        Condition = 9
        Status = 10
        Borrower = 11
    End Enum
End Module
