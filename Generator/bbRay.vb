Public Class bbRay 'polar coordinates
    Private Const IntersectionError As Double = 0.000000001

    Public x As Double = 0
    Public y As Double = 0
    Public angle As Double = 0

    Public Sub New(ByVal unitsX As Double, ByVal unitsY As Double, ByVal radiansAngle As Double)
        x = unitsX
        y = unitsY
        angle = radiansAngle
    End Sub

    Public Sub New(ByVal p1 As bbPoint, ByVal p2 As bbPoint)
        'Convert line segment p1----p2 into ray... p1---->
        x = p1.X
        y = p1.Y

        Dim temporary_Point As New bbPoint(p2.X - p1.X, p2.Y - p1.Y)
        angle = temporary_Point.Angle
    End Sub

    Public Function FindIntersection(ByVal OtherRay As bbRay) As bbPoint
        'This function finds the intersection between two rays without
        'having to solve the line equation using slopes and y-intersects.
        'Note that the complement to the ray is considered part of the ray
        '(i.e. a line)
        'It does this by using a series of simple transformations and
        'trig:
        '1) Translate the rays so that the first ray originates at (0, 0) the origin
        '2) Rotate the rays so that the first ray has an angle of 0
        '   This means that the whereever the second ray interesects the x-axis,
        '   it also intersects the first ray
        '3) Determine the x-intersect of the second ray using trig
        '4) Unrotate the system, and untranslate.
        '
        'It is true that this is a very slow way of computing the line
        'intersection; however, the advantage is that we never have to deal
        'with line segments, and we can more easily transform rays through
        'translation, rotation, and scaling than we could if we just had
        'a set of four points making up the two line segments.


        'First translate the rays so that the first one is at the origin
        Dim t1x As Double = x - OtherRay.x
        Dim t1y As Double = y - OtherRay.y
        Dim t1d As Double = Math.Sqrt(t1x * t1x + t1y * t1y)

        'Get the angle of t1 with respect to the origin
        Dim r1 As New bbRay(New bbPoint, New bbPoint(t1x, t1y))
        Dim t1a As Double = r1.angle

        Dim t2a = t1a - OtherRay.angle
        Dim t2d = t1d 'distance from origin stays the same
        Dim t2x = Math.Cos(t2a) * t2d
        Dim t2y = Math.Sin(t2a) * t2d
        Dim t2ray_angle = angle - OtherRay.angle

        'Dim intersection_y As Double = 0 (this line is unnecessary, 0 is implicit)
        Dim intersection_x As Double

        If Math.Abs(Math.Sin(t2ray_angle)) > IntersectionError Then
            intersection_x = t2x - t2y / Math.Tan(t2ray_angle)
        Else
            intersection_x = t2x
        End If

        Dim intersection_d As Double = Math.Abs(intersection_x)

        Dim t3x = Math.Cos(OtherRay.angle) * intersection_d + OtherRay.x
        Dim t3y = Math.Sin(OtherRay.angle) * intersection_d + OtherRay.y

        Return New bbPoint(t3x, t3y)
    End Function
End Class