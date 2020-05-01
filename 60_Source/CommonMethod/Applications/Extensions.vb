Imports System.Runtime.CompilerServices

Namespace Applications

    Friend Module Extensions

        <Extension>
        Public Function ToBooleanOrDefault(value As String, Optional defaultValue As Boolean = Nothing) As Boolean

            Dim result As Boolean

            If Boolean.TryParse(value, result) Then
                Return result
            Else
                Return defaultValue
            End If

        End Function

        <Extension>
        Public Function ToIntegerOrDefault(value As String) As Integer

            Dim result As Integer

            If Integer.TryParse(value, result) Then
                Return result
            Else
                ' default
                Return Nothing
            End If

        End Function

    End Module

End Namespace
