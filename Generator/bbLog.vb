Public Class bbLog
    Private ErrorLog As New Collection

    Public Sub LogError(ByRef ErrorText As String)
        ErrorLog.Add(ErrorText)
    End Sub

    Public Function ErrorCount() As Integer
        Return ErrorLog.Count()
    End Function

    Public Function ReadErrors(ByRef Index As Integer) As String
        'Return a specific error
        If Index >= ErrorLog.Count Then
            Return ""
            Exit Function
        End If
        Return CStr(CStr((Index + 1)) & CStr(ErrorLog.Item(Index - 1)))
    End Function

    Public Function ReadErrors() As String
        'Return all the errors
        ReadErrors = ""
        For i As Integer = 1 To ErrorCount()
            ReadErrors &= CStr(ErrorLog.Item(i) & Chr(13) & Chr(10))
        Next
        Return ReadErrors
    End Function
End Class