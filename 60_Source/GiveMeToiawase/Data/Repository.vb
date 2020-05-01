Imports System.Data.Common
Imports System.Data.SqlClient
Imports System.Data
Imports System.IO
Imports System.Text

Public Class Repository
    Private ReadOnly _connection As DbConnection
    Public Sub New(connection As DbConnection)
        _connection = connection
    End Sub
    Public Sub New(ConnectionString As String)
        Me.New(New SqlConnection(ConnectionString))
    End Sub
    Public Function BeginTrans() As IDbTransaction
        _connection.Open()
        Return _connection.BeginTransaction()
    End Function

    Public Sub RollBackTrans(trans As IDbTransaction)
        trans.Rollback()
        _connection.Close()
    End Sub

    Public Sub CommitTrans(trans As IDbTransaction)
        trans.Commit()
        _connection.Close()
    End Sub
    Public Function GetData(xSqlStr As String) As DataTable
        Try
            _connection.Open()
            Dim comm As New SqlCommand(xSqlStr, _connection)
            Dim reader As SqlDataReader = comm.ExecuteReader
            Dim dt As New DataTable
            dt.Load(reader)
            Return dt
        Catch ex As Exception
            Throw
        Finally
            _connection.Close()
        End Try
    End Function

    Public Sub excutedFolder(path As String, trans As IDbTransaction)
        If (_connection.State = ConnectionState.Closed) Then
            _connection.Open()
        End If
        Dim di As DirectoryInfo = New DirectoryInfo(path)
        Dim script As New StringBuilder
        For Each f In di.GetFiles("*.sql")
            script.AppendLine(File.ReadAllText(f.FullName))
        Next
        If script.Length > 0 Then
            Dim comm As New SqlCommand(script.ToString, _connection, trans)
            comm.ExecuteNonQuery()
        End If
    End Sub
End Class
