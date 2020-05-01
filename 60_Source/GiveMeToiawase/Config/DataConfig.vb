Public Class DataConfig

    Protected Const FILE_EXTEXNTION As String = "sql"
    Public Property TableName As String
    Public Property PGCodeColumn As String
    Public Property PGCodeSource As String
    Public Property PGCodeTargert As String
    Public Property Sort As New ArrayList
    Public Property Where As New SerializableDictionary(Of String, String)

    Public Function GetFileName()
        Return String.Format("{0}.{1}", TableName, FILE_EXTEXNTION)
    End Function
    Public Function GetSql()
        Dim sqlSelect As String = String.Empty
        Dim sqlWhere As String = String.Empty
        Dim sqlOrder As String = String.Empty

        sqlSelect = String.Format(" SELECT * FROM {0} ", Me.TableName)

        'If Not IsNothing(Where) AndAlso Where.Count > 0 Then
        '    sqlWhere = " WHERE 1= 1"
        '    For Each item In Where
        '        sqlWhere &= String.Format(" AND {0} = {1}", item.Key, item.Value)
        '    Next
        'End If

        Dim isDirty As Boolean = False
        If Not IsNothing(Sort) AndAlso Sort.Count > 0 Then
            sqlOrder = " ORDER BY "
            For Each item In Sort
                sqlOrder &= String.Format(" {0} {1} ", IIf(isDirty, ",", ""), item.ToString)
                isDirty = True
            Next
        End If

        Return String.Format("{0}{1}{2}", sqlSelect, sqlWhere, sqlOrder)

    End Function

End Class
