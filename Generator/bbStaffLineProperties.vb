Public Class bbStaffLineProperties
    'Change helpers
    Private RecentlyChanged As Boolean = False


    'Line Characteristics
    '--------------------

    'Line visibility
    Public Draw As New bbCascadingProperty           'boolean

    'End of line characteristics
    Public DrawFlatCap As New bbCascadingProperty    'boolean

    'Dash properties
    Public DrawDashedLine As New bbCascadingProperty 'boolean
    Public DashOffLength As New bbCascadingProperty  'double
    Public DashOnLength As New bbCascadingProperty   'double
    Public DashOffset As New bbCascadingProperty     'double

    'Line characteristics
    Public LineThickness As New bbCascadingProperty  'page units

    'Line color
    Public LineStrokeColor As New bbCascadingProperty      'bbColor

    Public Function SetProperties() As String
        Dim X As String = ""
        X &= PathCap(Not Me.DrawFlatCap.Value) & TokenSeperator()
        X &= PathJoin(True) & TokenSeperator()
        X &= PathWidth(Me.LineThickness.Value) & TokenSeperator()
        X &= PathDash(Me.DrawDashedLine.Value, Me.DashOnLength.Value, Me.DashOffLength.Value, Me.DashOffset.Value) & TokenSeperator()

        Dim lcObject As bbColor = LineStrokeColor.Value
        If Not IsNothing(lcObject) Then
            X &= PathColor(lcObject.r, lcObject.g, lcObject.b)
        End If

        Return X
    End Function

    Public Function Changed() As Boolean
        If RecentlyChanged Then
            RecentlyChanged = False
            Return True
        Else
            Return False
        End If
    End Function
    Public Shared Operator +(ByVal ShallowerObject As bbStaffLineProperties, ByVal DeeperObject As bbStaffLineProperties) As bbStaffLineProperties
        Dim x As New bbStaffLineProperties
        With DeeperObject
            If Not (IsNothing(.DashOffLength.Value) And _
                    IsNothing(.DashOffset.Value) And _
                    IsNothing(.DashOnLength.Value) And _
                    IsNothing(.Draw.Value) And _
                    IsNothing(.DrawDashedLine.Value) And _
                    IsNothing(.DrawFlatCap.Value) And _
                    IsNothing(.LineThickness.Value) And _
                    IsNothing(.LineStrokeColor.Value)) Then
                x.RecentlyChanged = True
            End If
        End With

        x.DashOffLength = ShallowerObject.DashOffLength + DeeperObject.DashOffLength
        x.DashOffset = ShallowerObject.DashOffset + DeeperObject.DashOffset
        x.DashOnLength = ShallowerObject.DashOnLength + DeeperObject.DashOnLength
        x.Draw = ShallowerObject.Draw + DeeperObject.Draw
        x.DrawDashedLine = ShallowerObject.DrawDashedLine + DeeperObject.DrawDashedLine
        x.DrawFlatCap = ShallowerObject.DrawFlatCap + DeeperObject.DrawFlatCap
        x.LineThickness = ShallowerObject.LineThickness + DeeperObject.LineThickness
        x.LineStrokeColor = ShallowerObject.LineStrokeColor + DeeperObject.LineStrokeColor
        Return x
    End Operator
End Class
