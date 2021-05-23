Public Class bbStaffProperties
    'The collection of all staff lines, with index 0 representing the path,
    'negative indexes referring to lines below and positive indexes referring
    'to lines above.
    Private Lines As New bbSmartCollection() 'Collection of bbStaffLineProperties

    'Attributes that apply to all of the lines
    Public SpaceHeight As New bbCascadingProperty

    Public Function LinesLBound() As Integer
        Return Lines.LBound()
    End Function

    Public Function LinesUBound() As Integer
        Return Lines.UBound()
    End Function

    Public Function LineCount() As Integer
        Return Lines.Count()
    End Function

    Public Property StaffLine(ByVal StaffLineIndex As Integer) As bbStaffLineProperties
        Get
            'Force data at Index to exist in case Get is called to obtain a reference
            'to internal members. This is a public method intended to help
            'outside commands get access to select information quickly...
            Dim x As bbStaffLineProperties = Lines(StaffLineIndex)
            If IsNothing(x) Then
                'We have to return something, and we have to make 
                'sure it gets added to the collection
                x = New bbStaffLineProperties
                Lines(StaffLineIndex) = x
            End If
            Return x
        End Get
        Set(ByVal value As bbStaffLineProperties)
            Lines(StaffLineIndex) = value
        End Set
    End Property
End Class
