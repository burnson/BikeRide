Module bbBasicSymbols
    Private Const degreesHalfNoteSlitRotation As Double = 20
    Private Const radiansHalfNoteSlitRotation As Double = (degreesHalfNoteSlitRotation / 360.0) * Math.PI * 2.0
    Private Const unitsHalfNoteSlitLength = 0.8
    Private Const unitsHalfNoteSlitThickness = 0.4

    Public Function GetNoteWidth(ByVal spaceHeight As Double) As Double
        Dim CosEllipse As Double
        Dim SinEllipse As Double
        Dim a As Double
        Dim b As Double
        Dim m As Double
        Dim n As Double

        n = spaceHeight / 2
        CosEllipse = Math.Cos(Globals.radiansNoteRotation) ^ 2
        SinEllipse = Math.Sin(Globals.radiansNoteRotation) ^ 2
        b = n / (Globals.ratioNoteWidthToNoteHeight * (1 - CosEllipse) + CosEllipse)
        a = Globals.ratioNoteWidthToNoteHeight * b
        m = a + (b - a) * SinEllipse
        Return m
    End Function

    Public Function QuarterNote( _
                                ByVal canvasStemOrigin As bbPoint, _
                                ByVal canvasStemWidth As Double, _
                                ByVal boolDrawLeftOfStem As Boolean, _
                                ByVal radiansOrientation As Double, _
                                ByVal spaceHeight As Double, _
                                ByVal percentageNoteSize As Double, _
                                ByRef colorNotehead As bbColor, _
                                ByRef canvasVerticalStemTrim As Double _
                                ) As String

        DrawQuarterNoteTime.StartClock()
        Dim Operators As String = ""
        Operators &= SaveCTM() & TokenSeperator()
        Operators &= CTMTranslate(canvasStemOrigin.X, canvasStemOrigin.Y) & TokenSeperator()
        Operators &= CTMRotate(radiansOrientation) & TokenSeperator()


        '-----------------------------
        'Derivation:
        'w = a + (b - a)|sin theta|
        'h = a + (b - a)|cos theta|
        'a = b * r
        'h = (br)+(b-br)(cos)
        'h = (br)+cos()[b]-cos()[br]
        ' = br*(1-cos())+cos()(b)
        ' = b[r(1-cos())+cos()]
        'b = h / [r(1-cos())+cos()]
        '-----------------------------

        Dim CosEllipse As Double
        Dim SinEllipse As Double
        Dim a As Double
        Dim b As Double
        Dim m As Double
        Dim n As Double
        Dim canvasNoteCenter As New bbPoint
        Dim ThetaEllipse As Double
        Dim rEllipse As Double
        Dim xFinal As Double, yFinal As Double
        Dim xNotRotated As Double, yNotRotated As Double
        Dim Ecc As Double

        n = spaceHeight * percentageNoteSize / 2
        CosEllipse = Math.Cos(Globals.radiansNoteRotation) ^ 2
        SinEllipse = Math.Sin(Globals.radiansNoteRotation) ^ 2
        b = n / (Globals.ratioNoteWidthToNoteHeight * (1 - CosEllipse) + CosEllipse)
        a = Globals.ratioNoteWidthToNoteHeight * b
        m = a + (b - a) * SinEllipse

        ThetaEllipse = Math.Atan(b * Math.Tan(-Globals.radiansNoteRotation) / a)
        Ecc = Math.Sqrt(1 - (b ^ 2) / (a ^ 2))
        rEllipse = b / Math.Sqrt(1 - (Ecc ^ 2) * (Math.Cos(ThetaEllipse) ^ 2))


        xNotRotated = rEllipse * Math.Cos(ThetaEllipse)
        yNotRotated = rEllipse * Math.Sin(ThetaEllipse)

        xFinal = xNotRotated * Math.Cos(Globals.radiansNoteRotation) - yNotRotated * Math.Sin(Globals.radiansNoteRotation)
        yFinal = xNotRotated * Math.Sin(Globals.radiansNoteRotation) + yNotRotated * Math.Cos(Globals.radiansNoteRotation)

        'Note: 'xFinal' should be equal to 'm' -- different equations, same result

        If boolDrawLeftOfStem Then
            canvasNoteCenter.X = -xFinal + canvasStemWidth / 2
            canvasNoteCenter.Y = 0
            canvasVerticalStemTrim = yFinal
        Else
            canvasNoteCenter.X = xFinal - canvasStemWidth / 2
            canvasNoteCenter.Y = 0
            canvasVerticalStemTrim = -yFinal
        End If



        Operators &= CTMTranslate(canvasNoteCenter.X, canvasNoteCenter.Y) & TokenSeperator()
        Operators &= CTMRotate(Globals.radiansNoteRotation) & TokenSeperator()
        Operators &= CTMScale(a * 2, b * 2) & TokenSeperator()
        Operators &= PathColor(colorNotehead.r, colorNotehead.g, colorNotehead.b) & TokenSeperator() 'just for now...
        Operators &= PathWidth(1) & TokenSeperator()
        Operators &= PathCap(True) & TokenSeperator()
        Operators &= PathStart(0, 0) & TokenSeperator()
        Operators &= PathContinue(0, 0) & TokenSeperator()
        Operators &= DrawPath() & TokenSeperator()

        Operators &= RevertCTM()
        DrawQuarterNoteTime.StopClock()
        Return Operators
    End Function
    Public Function HalfNote( _
                                ByVal canvasStemOrigin As bbPoint, _
                                ByVal canvasStemWidth As Double, _
                                ByVal boolDrawLeftOfStem As Boolean, _
                                ByVal radiansOrientation As Double, _
                                ByVal spaceHeight As Double, _
                                ByVal percentageNoteSize As Double, _
                                ByRef colorNotehead As bbColor, _
                                ByRef canvasVerticalStemTrim As Double _
                                ) As String

        Dim Operators As String = ""
        Operators &= SaveCTM() & TokenSeperator()
        Operators &= CTMTranslate(canvasStemOrigin.X, canvasStemOrigin.Y) & TokenSeperator()
        Operators &= CTMRotate(radiansOrientation) & TokenSeperator()


        '-----------------------------
        'Derivation:
        'w = a + (b - a)|sin theta|
        'h = a + (b - a)|cos theta|
        'a = b * r
        'h = (br)+(b-br)(cos)
        'h = (br)+cos()[b]-cos()[br]
        ' = br*(1-cos())+cos()(b)
        ' = b[r(1-cos())+cos()]
        'b = h / [r(1-cos())+cos()]
        '-----------------------------

        Dim CosEllipse As Double
        Dim SinEllipse As Double
        Dim a As Double
        Dim b As Double
        Dim m As Double
        Dim n As Double
        Dim canvasNoteCenter As New bbPoint
        Dim ThetaEllipse As Double
        Dim rEllipse As Double
        Dim xFinal As Double, yFinal As Double
        Dim xNotRotated As Double, yNotRotated As Double
        Dim Ecc As Double

        n = spaceHeight * percentageNoteSize / 2
        CosEllipse = Math.Cos(Globals.radiansNoteRotation) ^ 2
        SinEllipse = Math.Sin(Globals.radiansNoteRotation) ^ 2
        b = n / (Globals.ratioNoteWidthToNoteHeight * (1 - CosEllipse) + CosEllipse)
        a = Globals.ratioNoteWidthToNoteHeight * b
        m = a + (b - a) * SinEllipse

        ThetaEllipse = Math.Atan(b * Math.Tan(-Globals.radiansNoteRotation) / a)
        Ecc = Math.Sqrt(1 - (b ^ 2) / (a ^ 2))
        rEllipse = b / Math.Sqrt(1 - (Ecc ^ 2) * (Math.Cos(ThetaEllipse) ^ 2))


        xNotRotated = rEllipse * Math.Cos(ThetaEllipse)
        yNotRotated = rEllipse * Math.Sin(ThetaEllipse)

        xFinal = xNotRotated * Math.Cos(Globals.radiansNoteRotation) - yNotRotated * Math.Sin(Globals.radiansNoteRotation)
        yFinal = xNotRotated * Math.Sin(Globals.radiansNoteRotation) + yNotRotated * Math.Cos(Globals.radiansNoteRotation)

        'Note: 'xFinal' should be equal to 'm' -- different equations, same result

        If boolDrawLeftOfStem Then
            canvasNoteCenter.X = -xFinal + canvasStemWidth / 2
            canvasNoteCenter.Y = 0
            canvasVerticalStemTrim = yFinal
        Else
            canvasNoteCenter.X = xFinal - canvasStemWidth / 2
            canvasNoteCenter.Y = 0
            canvasVerticalStemTrim = -yFinal
        End If




        Operators &= CTMTranslate(canvasNoteCenter.X, canvasNoteCenter.Y) & TokenSeperator()
        Operators &= CTMRotate(Globals.radiansNoteRotation) & TokenSeperator()
        Operators &= CTMScale(a * 2, b * 2) & TokenSeperator()


        Operators &= PathColor(colorNotehead.r, colorNotehead.g, colorNotehead.b) & TokenSeperator() 'just for now...
        Operators &= PathCap(True) & TokenSeperator()
        Operators &= PathWidth(0.2) & TokenSeperator()





        Operators &= CTMRotate(radiansHalfNoteSlitRotation) & TokenSeperator()


        Operators &= PathStart(-2, -2) & TokenSeperator() 'clockwise
        Operators &= PathContinue(-2, 2) & TokenSeperator()
        Operators &= PathContinue(2, 2) & TokenSeperator()
        Operators &= PathContinue(2, -2) & TokenSeperator()
        Operators &= StartNewSubPath() & TokenSeperator()

        Dim p As Double = unitsHalfNoteSlitLength / 2
        Dim q As Double = unitsHalfNoteSlitThickness / 2
        Dim TotalPoints As Integer = 100
        For i As Integer = 0 To TotalPoints - 1
            Dim x As Double, y As Double
            Dim k As Double
            If i < TotalPoints / 2 Then
                k = p - q
            Else
                k = -p + q
            End If

            x = q * Math.Cos((i / TotalPoints - 0.25) * Math.PI * 2) + k
            y = q * Math.Sin((i / TotalPoints - 0.25) * Math.PI * 2)

            If i = 0 Then
                Operators &= PathStart(x, y) & TokenSeperator() 'counter-clockwise
            Else
                Operators &= PathContinue(x, y) & TokenSeperator()
            End If
        Next

        Operators &= DrawClippingPath(True) & TokenSeperator()
        Operators &= DoNotDrawPath() & TokenSeperator()

        Operators &= CTMRotate(-radiansHalfNoteSlitRotation) & TokenSeperator()

        Operators &= PathWidth(1) & TokenSeperator()
        Operators &= PathStart(0, 0) & TokenSeperator()
        Operators &= PathContinue(0, 0) & TokenSeperator()
        Operators &= DrawPath() & TokenSeperator()

        Operators &= RevertCTM()

        Return Operators

    End Function
End Module
