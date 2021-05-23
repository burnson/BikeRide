Public Class bbCascadingProperty
    Public Value As Object

    Public Sub New()
        Value = Nothing
    End Sub

    Public Sub New(ByVal InitialValue As Object)
        Value = InitialValue
    End Sub

    Public Shared Operator +(ByVal ShallowerObject As bbCascadingProperty, ByVal DeeperObject As bbCascadingProperty) As bbCascadingProperty

        'Basic cascading property rules:
        '                 a + b = b
        '           Nothing + b = b
        '           a + Nothing = a
        '     Nothing + Nothing = Nothing
        '
        'Essentially: assume b, unless it has no information

        If IsNothing(ShallowerObject.Value) Or Not IsNothing(DeeperObject.Value) Then
            Return DeeperObject
        Else
            Return ShallowerObject
        End If
    End Operator

    Public Sub Reset()
        Value = Nothing
    End Sub

    Public Sub Reset(ByVal NewValue As Object)
        Value = NewValue
    End Sub
End Class
