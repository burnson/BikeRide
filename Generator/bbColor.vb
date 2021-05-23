Public Class bbColor
    Public r As Double
    Public g As Double
    Public b As Double

    Public Sub New()
        r = 0
        g = 0
        b = 0
    End Sub

    Public Sub New(ByVal Red As Double, ByVal Green As Double, ByVal Blue As Double)
        r = Red
        g = Green
        b = Blue
    End Sub

    Public Sub New(ByVal SystemColor As System.Drawing.Color)
        r = CDbl(SystemColor.R) / 255.0
        g = CDbl(SystemColor.G) / 255.0
        b = CDbl(SystemColor.B) / 255.0
    End Sub
End Class
