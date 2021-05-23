Public Class bbPDFObject
    Private ObjectLines As New Collection

    Public Sub Write(ByRef str As String)
        ObjectLines.Add(str)
    End Sub

    Public Function LineCount() As Integer
        Return ObjectLines.Count
    End Function

    Public Function GetLine(ByVal index As Integer) As String
        Return ObjectLines(index)
    End Function
End Class