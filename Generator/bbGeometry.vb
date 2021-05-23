Module bbGeometry
    Public Function FindDisplacement(ByRef Point1 As bbPoint, ByRef Point2 As bbPoint, ByVal Displacement As Double) As bbPoint
        Dim Ray1 As New bbRay(Point1, Point2)

        Ray1.x += Math.Cos(Ray1.angle + Math.PI / 2) * Displacement
        Ray1.y += Math.Sin(Ray1.angle + Math.PI / 2) * Displacement

        Return New bbPoint(Ray1.x, Ray1.y)
    End Function


    Public Function FindIntersectionAfterDisplacement(ByRef Point1 As bbPoint, ByRef Point2 As bbPoint, ByRef Point3 As bbPoint, ByVal Segment1Displacement As Double, ByVal Segment2Displacement As Double) As bbPoint
        Dim Ray1 As New bbRay(Point1, Point2)
        Dim Ray2 As New bbRay(Point2, Point3)

        Ray1.x += Math.Cos(Ray1.angle + Math.PI / 2) * Segment1Displacement
        Ray1.y += Math.Sin(Ray1.angle + Math.PI / 2) * Segment1Displacement

        Ray2.x += Math.Cos(Ray2.angle + Math.PI / 2) * Segment2Displacement
        Ray2.y += Math.Sin(Ray2.angle + Math.PI / 2) * Segment2Displacement

        Return Ray1.FindIntersection(Ray2)
    End Function
End Module
