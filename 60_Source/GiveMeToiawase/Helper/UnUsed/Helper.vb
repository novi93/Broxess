Imports System.Data.SqlClient
Imports System.Data

Public Class Helper
    Protected _repo As Repository

    Public Sub New(repo As Repository)
        _repo = repo
    End Sub

    Public Sub GetData(configs As List(Of DataConfig))

        Dim xSqlStr As String = String.Empty

        'For Each config In configs
        '    xSqlStr = config.GetSql()
        '    Dim dt As DataTable = _repo.GetData(xSqlStr)
        '    SqlGenerator.Save2File(SqlGenerator.GenerateScript(dt), "D:\1.sql")
        'Next

    End Sub
End Class
