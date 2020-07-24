Imports System.Runtime.CompilerServices


Module StringExtension

    ''' <summary>
    ''' 「`」→ 「''」
    ''' </summary>
    ''' <param name="source"></param>
    <Extension()>
    Public Function MiniSlashTo2Qoute(source As String)
        Return source.Replace("`"c, "''")
    End Function
End Module
