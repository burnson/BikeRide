'The following class represents a point
'and uses both polar and rectangular coordinates
'in tandem


Public Class bbPoint
    Private rect_X As Double
    Private rect_Y As Double
    Private polar_Angle As Double
    Private polar_Radius As Double

    Private Sub PolarToRectangular(ByVal nAngle As Double, ByVal nRadius As Double, ByRef nX As Double, ByRef nY As Double)
        nX = Math.Cos(nAngle) * nRadius
        nY = Math.Sin(nAngle) * nRadius
    End Sub

    Private Sub RectangularToPolar(ByVal nX As Double, ByVal nY As Double, ByRef nAngle As Double, ByRef nRadius As Double)
        nRadius = Math.Sqrt(nX * nX + nY * nY)

        'This computes the angle -- in order to avoid dealing with
        'the asymptotes, if |Y| > |X|  then the coordinates are
        'switched and the angle is later corrected by adding pi/2
        '
        'This method allows the angle to always be accurately computed
        'even in circumstances where the internal representation
        'might have a smaller number of bits

        If nRadius = 0 Then
            nAngle = 0
        ElseIf Math.Abs(nY) < Math.Abs(nX) Then
            If nX > 0 Then
                nAngle = Math.Atan(nY / nX)
            Else
                nAngle = Math.Atan(nY / nX) + Math.PI
            End If
        Else
            If nY > 0 Then
                nAngle = -Math.Atan(nX / nY) + Math.PI / 2
            Else
                nAngle = -Math.Atan(nX / nY) + Math.PI * 3 / 2
            End If
        End If
    End Sub

    Public Sub New(ByVal nX As Double, ByVal nY As Double)
        'Unfortunately can't overload for polar coordinates since it would
        'have the same number and type of arguments... oh well. The properties
        'allow easy access to polar coordinates anyway...
        rect_X = nX
        rect_Y = nY
        RectangularToPolar(rect_X, rect_Y, polar_Angle, polar_Radius)
    End Sub

    Public Sub New()
        rect_X = 0
        rect_Y = 0
        polar_Angle = 0
        polar_Radius = 0
    End Sub

    Public Property X() As Double
        Get
            Return rect_X
        End Get
        Set(ByVal value As Double)
            rect_X = value
            RectangularToPolar(rect_X, rect_Y, polar_Angle, polar_Radius)
        End Set
    End Property

    Public Property Y() As Double
        Get
            Return rect_Y
        End Get
        Set(ByVal value As Double)
            rect_Y = value
            RectangularToPolar(rect_X, rect_Y, polar_Angle, polar_Radius)
        End Set
    End Property

    Public Property Angle() As Double
        Get
            Return polar_Angle
        End Get
        Set(ByVal value As Double)
            polar_Angle = value
            PolarToRectangular(polar_Angle, polar_Radius, rect_X, rect_Y)
        End Set
    End Property

    Public Property Radius() As Double
        Get
            Return polar_Radius
        End Get
        Set(ByVal value As Double)
            polar_Radius = value
            PolarToRectangular(polar_Angle, polar_Radius, rect_X, rect_Y)
        End Set
    End Property
End Class
