Public Class bbSmartCollection
    'The smart collection basically allows you to plot
    'data at integer locations without presuming that
    'the elements in between have data.
    '
    'For example:
    'Dim x As New bbSmartCollection
    'x(0)="this is the first index"
    'x(2)="this is the third index"
    '
    'At this point x(1) would still be Nothing and there
    'would be no space allocated for it either. A default
    'value can also be specified in the constructor so
    'that something other than Nothing can be returned
    'in these cases.

    Private ObjectList As New Collection
    Private IndexList As New Collection

    Private DefaultValue As Object

    Private MinIndex As Integer = 0
    Private MaxIndex As Integer = -1 'This is so that Count() works properly with no elements...

    Public Sub New()
        DefaultValue = Nothing
    End Sub

    'Public Sub New(ByVal NoElementDefaultValue As Object)
    '    DefaultValue = NoElementDefaultValue
    'End Sub

    Default Public Property Value(ByVal Index As Integer) As Object
        Get
            'Looks up the value at the given index... if no value
            'is there it returns the "DefaultValue" specified by
            'the constructor, which is initially Nothing.
            If IndexList.Count > 0 Then
                With IndexList
                    Dim i As Integer = .Item(1), j As Integer = .Item(.Count)
                    If j - i = .Count - 1 Then
                        'Guess it as linear
                        Dim t As Integer = 1 + (Index - i)
                        If t <= .Count Then
                            Dim y As Object = IndexList(t)
                            Dim z As Integer = CLng(y)
                            If z = Index Then
                                Return ObjectList(t)
                            End If
                        End If
                    End If
                End With

                For x As Integer = 1 To IndexList.Count
                    Dim y As Object = IndexList(x)
                    Dim z As Integer = CLng(y)
                    If z = Index Then
                        Return ObjectList(x)
                    End If
                Next
            End If
            Return DefaultValue
        End Get
        Set(ByVal value As Object)
            'Try to find the index, and replace the object there.
            'Otherwise add a new element
            For x As Integer = 1 To IndexList.Count
                Dim y As Object = IndexList(x)
                Dim z As Integer = CLng(y)
                If z = Index Then
                    Dim AfterObject As Object = ObjectList.Item(x)
                    ObjectList.Add(value, After:=x)
                    ObjectList.Remove(x)
                    Return
                End If
            Next
            ObjectList.Add(value)
            IndexList.Add(Index)
            If ObjectList.Count = 1 Then
                MinIndex = Index
                MaxIndex = Index
            Else
                If Index < MinIndex Then
                    MinIndex = Index
                End If
                If Index > MaxIndex Then
                    MaxIndex = Index
                End If
            End If
        End Set
    End Property

    Public Function LBound() As Integer
        Return MinIndex
    End Function

    Public Function UBound() As Integer
        Return MaxIndex
    End Function

    Public Function Count() As Integer
        Return UBound() - LBound() + 1
    End Function

    Public Function FindNearestNeighbor(ByVal Index As Integer) As Integer
        If IndexList.Count = 0 Then Return 0 'Inevitably possible if there are no elements

        'See if we can find the real thing first
        For x As Integer = 1 To IndexList.Count
            Dim y As Object = IndexList(x)
            Dim z As Integer = CLng(y)
            If z = Index Then
                Return z
            End If
        Next

        'We couldn't find the index, so we have to find the closest one
        'sweeping positive to negative
        If Index < LBound() Then
            Index = LBound()
        ElseIf Index > UBound() Then
            Index = UBound()
        Else
            'This one is tricky! We'll have to do a min distance
            'algorithm to find the best one. Can't sweep through
            'all of them one by one in a linear search because there
            'could potentially be quadrillions in between!

            Dim MaxIndex As Integer = LBound()
            For i As Integer = 1 To IndexList.Count
                If Not Index <= IndexList(i) Then
                    'We must ignore anything above the index, since
                    'we are reading from negative to positive
                    If IndexList(i) > MaxIndex Then
                        MaxIndex = IndexList(i)
                    End If
                End If
            Next i
            'So now MaxIndex contains the most positive index that is
            'less than Index. Phew...
            Return MaxIndex
        End If
    End Function
End Class
