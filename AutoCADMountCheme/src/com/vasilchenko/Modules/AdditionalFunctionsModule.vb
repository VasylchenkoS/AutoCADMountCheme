Imports System.Text.RegularExpressions

Namespace com.vasilchenko.Modules
    Module AdditionalFunctionsModule

        Public Function GetLastNumericFromString(s As String) As Double
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                s = Replace(matches(matches.Count - 1).Value, ".", ",")
                Return Math.Abs(Double.Parse(s))
            Else
                Return -1
            End If
        End Function

        Public Function GetFirstNumericFromString(s As String) As Double
            If IsNothing(s) Then Return -1
            Dim rgx As New Regex("-?\d*\.?\d+", RegexOptions.IgnoreCase)
            Dim matches As MatchCollection = rgx.Matches(s)
            If matches.Count > 0 Then
                s = Replace(matches(0).Value, ".", ",")
                Return Math.Abs(Double.Parse(s))
            Else
                Return -1
            End If
        End Function

    End Module
End Namespace
