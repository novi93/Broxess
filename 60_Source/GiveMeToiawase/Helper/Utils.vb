Public Class Utils

    Protected Const CONNECT_STRING = "Data Source={0};Initial Catalog={1};User ID={2};Password={3};"

    Public Shared Function createConnectString(server As String,
                                               isLocal As Boolean,
                                     user As String,
                                     pass As String,
                                     Optional dbName As String = Nothing) As String
        If String.IsNullOrWhiteSpace(dbName) Then
            dbName = "master"
        End If
        Return String.Format(CONNECT_STRING, server, dbName, user, pass)
    End Function
    Shared Sub RunCommandCom(command As String, arguments As String, permanent As Boolean)
        Dim p As Process = New Process()
        Dim pi As ProcessStartInfo = New ProcessStartInfo()
        pi.Arguments = " " + If(permanent = True, "/K", "/C") + " " + command + " " + arguments
        pi.FileName = "cmd.exe"
        p.StartInfo = pi
        p.Start()
    End Sub
End Class
