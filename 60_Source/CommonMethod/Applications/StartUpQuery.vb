Imports System.Linq
Imports System.Net
Imports System.Reflection
Imports System.Web

Namespace Applications

    ''' <summary>
    ''' アプリケーションの起動に使用されるクエリーです。
    ''' </summary>
    Public Class StartUpQuery

        ''' <summary>
        ''' アプリケーションの起動に必要なパラメーターで新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="rootEndPoint"></param>
        ''' <param name="accessToken"></param>
        ''' <param name="programCode"></param>
        ''' <param name="startUpKey"></param>
        Public Sub New(rootEndPoint As String, accessToken As String, programCode As Integer, startUpKey As String)

            Me.RootEndPoint = rootEndPoint
            Me.AccessToken = accessToken
            Me.ProgramCode = programCode
            Me.StartUpKey = startUpKey

            ThrowIfInvalid()

        End Sub

        ''' <summary>
        ''' クエリー文字列で新しいインスタンスを初期化します。
        ''' </summary>
        ''' <param name="queryString"></param>
        Public Sub New(queryString As String)

            Dim parameters = queryString.TrimStart("?"c).Split("&"c) _
            .Select(Function(x) x.Split("="c)) _
            .Select(Function(x) New KeyValuePair(Of String, String)(x(0), HttpUtility.UrlDecode(x(1)))) _
            .ToArray()

            For Each p In parameters
                If String.Equals(p.Key, "RootEndPoint", StringComparison.OrdinalIgnoreCase) Then
                    RootEndPoint = p.Value

                ElseIf String.Equals(p.Key, "AccessToken", StringComparison.OrdinalIgnoreCase) Then
                    AccessToken = p.Value

                ElseIf String.Equals(p.Key, "ProgramCode", StringComparison.OrdinalIgnoreCase) Then
                    ProgramCode = p.Value.ToIntegerOrDefault()

                ElseIf String.Equals(p.Key, "StartUpKey", StringComparison.OrdinalIgnoreCase) Then
                    StartUpKey = p.Value

                ElseIf String.Equals(p.Key, "FlowOption", StringComparison.OrdinalIgnoreCase) Then
                    FlowOption = p.Value

                ElseIf String.Equals(p.Key, "Timeout", StringComparison.OrdinalIgnoreCase) Then
                    Timeout = p.Value.ToIntegerOrDefault()

                ElseIf String.Equals(p.Key, "Expect100Continue", StringComparison.OrdinalIgnoreCase) Then
                    Expect100Continue = p.Value.ToBooleanOrDefault(ServicePointManager.Expect100Continue)

                ElseIf String.Equals(p.Key, "DebugAppPath", StringComparison.OrdinalIgnoreCase) Then
                    DebugAppPath = p.Value

                End If
            Next

            ThrowIfInvalid()

        End Sub

        Public ReadOnly Property RootEndPoint As String
            Get
                Return String.Empty
            End Get
        End Property
        Public ReadOnly Property AccessToken As String
            Get
                Return String.Empty
            End Get
        End Property
        Public ReadOnly Property ProgramCode As Integer
            Get
                Return 0
            End Get
        End Property
        Public ReadOnly Property StartUpKey As String
            Get
                Return String.Empty
            End Get
        End Property

        Public Property FlowOption As String

        Public Property Timeout As Integer

        Public Property Expect100Continue As Boolean

        Public Property DebugAppPath As String

        Private Sub ThrowIfInvalid()

            Dim invalidParameters = New Dictionary(Of String, Boolean)() From
            {
                {"RootEndPoint", String.IsNullOrWhiteSpace(RootEndPoint)},
                {"AccessToken", String.IsNullOrWhiteSpace(AccessToken)},
                {"ProgramCode", ProgramCode = 0},
                {"StartUpKey", String.IsNullOrWhiteSpace(StartUpKey)}
            } _
            .Where(Function(x) x.Value) _
            .Select(Function(x) x.Key)

            If invalidParameters.Any() Then
                Throw New ArgumentException("起動に必要なクエリーが不足しています。: " + String.Join(", ", invalidParameters))
            End If

        End Sub

        Private Shared ReadOnly Properties As PropertyInfo() =
            GetType(StartUpQuery).GetProperties() _
                .Where(Function(x) x.CanRead) _
                .ToArray()

        Public Function ToQueryString() As String

            Return Properties _
                .Select(Function(x) New KeyValuePair(Of String, String)(x.Name, HttpUtility.UrlEncode(Convert.ToString(x.GetValue(Me))))) _
                .Where(Function(x) Not String.IsNullOrWhiteSpace(x.Value)) _
                .Select(Function(x) $"{x.Key}={x.Value}") _
                .Aggregate(Function(x, y) $"{x}&{y}")

        End Function

    End Class

End Namespace
