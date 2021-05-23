Module bbOperators
    Public Function CSpace() As String
        Return " "
    End Function

    Public Function TokenSeperator() As String
        Return CSpace()
    End Function

    Public Function PDFObjectStringID(ByVal Index As Integer) As String
        Return Index & CSpace() & "0" & CSpace() & "R"
    End Function

    Public Function CReal(ByVal x As Double) As String
        CReal = Format(x, "0.0##################") 'was 23
    End Function

    Public Function CTM(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByVal e As Double, ByVal f As Double) As String
        Return CReal(a) & CSpace() & CReal(b) & CSpace() & CReal(c) & CSpace() & CReal(d) & CSpace() & CReal(e) & CSpace() & CReal(f) & CSpace() & "cm" & TokenSeperator()
    End Function

    Public Function TextCTM(ByVal a As Double, ByVal b As Double, ByVal c As Double, ByVal d As Double, ByVal e As Double, ByVal f As Double) As String
        Return CReal(a) & CSpace() & CReal(b) & CSpace() & CReal(c) & CSpace() & CReal(d) & CSpace() & CReal(e) & CSpace() & CReal(f) & CSpace() & "Tm" & TokenSeperator()
    End Function

    Public Function TextCTMRotate(ByVal radiansAngle As Double) As String
        Dim i As Double = Math.Cos(radiansAngle)
        Dim j As Double = Math.Sin(radiansAngle)
        Return CTM(i, j, -j, i, 0, 0)
    End Function

    Public Function CTMTranslate(ByVal x As Double, ByVal y As Double) As String
        Return CTM(1, 0, 0, 1, x, y)
    End Function

    Public Function CTMScale(ByVal scaleX As Double, ByVal scaleY As Double) As String
        Return CTM(scaleX, 0, 0, scaleY, 0, 0)
    End Function

    Public Function CTMRotate(ByVal radiansAngle As Double) As String
        Dim i As Double = Math.Cos(radiansAngle)
        Dim j As Double = Math.Sin(radiansAngle)
        Return CTM(i, j, -j, i, 0, 0)
    End Function

    Public Function SaveCTM() As String
        Return "q" & TokenSeperator()
    End Function

    Public Function RevertCTM() As String
        Return "Q" & TokenSeperator()
    End Function

    Public Function PathStart(ByVal x As Double, ByVal y As Double) As String
        Return CReal(x) & CSpace() & CReal(y) & CSpace() & "m" & TokenSeperator()
    End Function

    Public Function PathContinue(ByVal x As Double, ByVal y As Double) As String
        Return CReal(x) & CSpace() & CReal(y) & CSpace() & "l" & TokenSeperator()
    End Function

    Public Function StartNewSubPath() As String
        Return "h" & TokenSeperator()
    End Function

    Public Function DrawPath() As String
        Return "S" & TokenSeperator()
    End Function

    Public Function DrawPathClosed() As String
        Return "s" & TokenSeperator()
    End Function

    Public Function DoNotDrawPath() As String
        Return "n" & TokenSeperator()
    End Function

    Public Function DrawClippingPath(ByVal Winding As Boolean) As String
        If Winding Then
            Return "W*" & TokenSeperator()
        Else
            Return "W" & TokenSeperator()
        End If
    End Function

    Public Function PathCap(ByVal Rounded As Boolean) As String
        If Rounded Then
            Return "1" & CSpace() & "J" & TokenSeperator()
        Else
            Return "0" & CSpace() & "J" & TokenSeperator()
        End If
    End Function

    Public Function PathDash(ByVal UseDash As Boolean, ByVal OnLength As Double, ByVal OffLength As Double, ByVal Offset As Double)
        OnLength = Math.Abs(OnLength)
        OffLength = Math.Abs(OffLength)
        If OnLength + OffLength = 0 Then
            UseDash = False
        Else
            Offset = (Offset Mod (OnLength + OffLength))
        End If
        If Not UseDash Then
            Return "[] 0 d" & TokenSeperator()
        Else
            Return "[" & CReal(OnLength) & CSpace() & CReal(OffLength) & "]" & CSpace() & Offset & CSpace() & "d" & TokenSeperator()
        End If
    End Function

    Public Function PathJoin(ByVal Rounded As Boolean) As String
        If Rounded Then
            Return "1" & CSpace() & "j" & TokenSeperator()
        Else
            Return "0" & CSpace() & "j" & TokenSeperator()
        End If
    End Function

    Public Function PathWidth(ByVal Width As Double) As String
        Width = Math.Abs(Width)
        Return CReal(Width) & CSpace() & "w" & TokenSeperator()
    End Function

    Public Function PathColor(ByVal R As Double, ByVal G As Double, ByVal B As Double) As String
        R = Math.Max(0, R)
        G = Math.Max(0, G)
        B = Math.Max(0, B)

        R = Math.Min(1, R)
        G = Math.Min(1, G)
        B = Math.Min(1, B)

        Return CReal(R) & CSpace() & CReal(G) & CSpace() & CReal(B) & CSpace() & "RG" & TokenSeperator()
    End Function

    Public Function DrawPoint(ByVal X As Double, ByVal Y As Double, ByVal Radius As Double, ByRef Color As bbColor) As String
        Dim stream As String = ""
        stream &= SaveCTM()
        stream &= PathCap(True)
        stream &= PathWidth(Radius)
        stream &= PathColor(Color.r, Color.g, Color.b)
        stream &= PathStart(X, Y)
        stream &= PathContinue(X, Y)
        stream &= DrawPath()
        stream &= RevertCTM()
        Return stream
    End Function

    Public Function DrawFilledRegion() As String
        Return "f" & TokenSeperator()
    End Function

    Public Function DrawPoint(ByVal X As Double, ByVal Y As Double, ByVal Radius As Double) As String
        Return DrawPoint(X, Y, Radius, New bbColor(0, 0, 0))
    End Function

    Public Function DrawPoint(ByVal Point As bbPoint, ByVal Radius As Double) As String
        Return DrawPoint(Point.X, Point.Y, Radius)
    End Function

    Public Function DrawLine(ByRef p1 As bbPoint, ByRef p2 As bbPoint, ByVal LineWidth As Double) As String
        Return DrawLine(p1, p2, LineWidth, New bbColor(0, 0, 0))
    End Function

    Public Function DrawLine(ByRef p1 As bbPoint, ByRef p2 As bbPoint, ByVal LineWidth As Double, ByVal Color As bbColor) As String
        Dim stream As String = ""
        stream &= SaveCTM()
        stream &= PathWidth(LineWidth)
        stream &= PathColor(Color.r, Color.g, Color.b)
        stream &= PathStart(p1.X, p1.Y)
        stream &= PathContinue(p2.X, p2.Y)
        stream &= DrawPath()
        stream &= RevertCTM()
        Return stream
    End Function

    Public Function DrawString(ByVal Font As bbFont, ByVal FontSize As Double, ByVal Orientation As bbRay, ByVal Text As String)
        Dim x As String = ""
        FontSize *= Font.Scale()
        x &= SaveCTM()
        x &= CTMTranslate(Orientation.x, Orientation.y)
        x &= CTMRotate(Orientation.angle) & TokenSeperator()
        x &= "BT" & TokenSeperator()
        x &= Font.GetPDFID() & CSpace() & CReal(FontSize) & CSpace() & "Tf" & TokenSeperator()
        x &= CReal(0) & CSpace() & CReal(0) & CSpace() & "Td" & TokenSeperator()
        x &= "(" & Text & ") Tj" & TokenSeperator()
        x &= "ET" & TokenSeperator()
        x &= RevertCTM()
        Return x
    End Function

    Public Function DrawSquare(ByVal Center As bbRay, ByVal Width As Double, ByVal LineWidth As Double) As String
        Dim tl As New bbPoint, tr As New bbPoint, bl As New bbPoint, br As New bbPoint
        Dim x As String = ""
        Width = Width / 2.0 * Math.Sqrt(2)
        Dim CosX As Double = Math.Cos(Center.angle + Math.PI / 4) * Width
        Dim SinY As Double = Math.Sin(Center.angle + Math.PI / 4) * Width


        tl.X = Center.x - SinY
        tl.Y = Center.y + CosX

        tr.X = Center.x + CosX
        tr.Y = Center.y + SinY

        bl.X = Center.x - CosX
        bl.Y = Center.y - SinY

        br.X = Center.x + SinY
        br.Y = Center.y - CosX

        x &= SaveCTM()
        x &= PathWidth(LineWidth)
        x &= PathJoin(True)
        x &= PathCap(False)
        x &= PathStart(tl.X, tl.Y)
        x &= PathContinue(tr.X, tr.Y)
        x &= PathContinue(br.X, br.Y)
        x &= PathContinue(bl.X, bl.Y)
        x &= DrawPathClosed()
        x &= RevertCTM()
        Return x
    End Function
End Module
