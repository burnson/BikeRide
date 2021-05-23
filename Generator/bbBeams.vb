Module bbBeams
    Public Sub DrawHairpin(ByRef Staff As bbStaff, ByVal IsCresc As Boolean, ByVal unitsLeftX As Double, ByVal spacesLeftY As Double, ByVal unitsRightX As Double, ByVal spacesRightY As Double, ByVal spacesHeight As Double, ByVal Thickness As Double, ByVal Interpolations As Integer, ByVal Color As bbColor)
        Dim Page As bbPDFPage = Nothing
        Dim PinchPoint As bbRay = Nothing
        Dim UpperFork As bbRay = Nothing
        Dim LowerFork As bbRay = Nothing
        Dim spaceheight1 As Double, spaceheight2 As Double, spaceheight3 As Double

        Staff.GetParent(Page)

        If IsCresc Then
            DrawBeam(Staff, unitsLeftX, spacesLeftY, 0, unitsRightX, spacesRightY + spacesHeight / 2, 0, Thickness, Thickness, Interpolations * 10, Color)
            DrawBeam(Staff, unitsLeftX, spacesLeftY, 0, unitsRightX, spacesRightY - spacesHeight / 2, 0, Thickness, Thickness, Interpolations * 10, Color)

            Staff.FindPlottingPoint(unitsLeftX, spacesLeftY, PinchPoint, spaceheight1)
            Staff.FindPlottingPoint(unitsRightX, spacesRightY + spacesHeight / 2.0, UpperFork, spaceheight2)
            Staff.FindPlottingPoint(unitsRightX, spacesRightY - spacesHeight / 2.0, LowerFork, spaceheight3)

        Else
            DrawBeam(Staff, unitsLeftX, spacesLeftY + spacesHeight / 2, 0, unitsRightX, spacesRightY, 0, Thickness, Thickness, Interpolations * 10, Color)
            DrawBeam(Staff, unitsLeftX, spacesLeftY - spacesHeight / 2, 0, unitsRightX, spacesRightY, 0, Thickness, Thickness, Interpolations * 10, Color)

            'bug'
            Staff.FindPlottingPoint(unitsRightX, spacesRightY, PinchPoint, spaceheight1)
            Staff.FindPlottingPoint(unitsLeftX, spacesLeftY + spacesHeight / 2.0, UpperFork, spaceheight2)
            Staff.FindPlottingPoint(unitsLeftX, spacesLeftY - spacesHeight / 2.0, LowerFork, spaceheight3)
        End If

        Page.AddLineToStream(DrawPoint(PinchPoint.x, PinchPoint.y, spaceheight1 * Globals.ratioStaffLineToSpaceHeight, Color))
        Page.AddLineToStream(DrawPoint(UpperFork.x, UpperFork.y, spaceheight1 * Globals.ratioStaffLineToSpaceHeight, Color))
        Page.AddLineToStream(DrawPoint(LowerFork.x, LowerFork.y, spaceheight1 * Globals.ratioStaffLineToSpaceHeight, Color))
    End Sub

    Public Sub DrawBeam(ByRef Staff As bbStaff, _
 _
                        ByVal unitsLeftStemX As Double, _
                        ByVal spacesLeftStemY As Double, _
                        ByVal LeftStemWidth As Double, _
 _
                        ByVal unitsRightStemX As Double, _
                        ByVal spacesRightStemY As Double, _
                        ByVal RightStemWidth As Double, _
 _
                        ByVal LeftBeamThickness As Double, _
                        ByVal RightBeamThickness As Double, _
 _
                        ByVal Interpolations As Integer, _
 _
                        ByVal Color As bbColor)

        Dim BeamTopPoints As New Collection
        Dim BeamBottomPoints As New Collection
        Dim Page As bbPDFPage = Nothing
        Staff.GetParent(Page)

        'Get the four corners of the beam
        '--------------------------------
        'To draw the beam, we will assume that whereever the stem terminates
        'is considered the middle of the beam. This way, even if the beam
        'is angled, it will completely cover the end of the stem.

        Dim TopLeftPoint As New bbPoint
        Dim BottomLeftPoint As New bbPoint
        Dim TopRightPoint As New bbPoint
        Dim BottomRightPoint As New bbPoint

        Dim LeftStemTop As bbRay = Nothing
        Dim RightStemTop As bbRay = Nothing
        Dim LeftSpaceHeight As Double = 0.0
        Dim RightSpaceHeight As Double = 0.0

        Staff.FindPlottingPoint(unitsLeftStemX, spacesLeftStemY, LeftStemTop, LeftSpaceHeight)
        Staff.FindPlottingPoint(unitsRightStemX, spacesRightStemY, RightStemTop, RightSpaceHeight)

    
        TopLeftPoint.X = LeftStemTop.x + Math.Cos(LeftStemTop.angle + Math.PI) * LeftStemWidth / 2.0 _
        + Math.Cos(LeftStemTop.angle + Math.PI / 2) * LeftBeamThickness * LeftSpaceHeight / 2.0
        TopLeftPoint.Y = LeftStemTop.y + Math.Sin(LeftStemTop.angle + Math.PI) * LeftStemWidth / 2.0 _
        + Math.Sin(LeftStemTop.angle + Math.PI / 2) * LeftBeamThickness * LeftSpaceHeight / 2.0

        BottomLeftPoint.X = LeftStemTop.x + Math.Cos(LeftStemTop.angle + Math.PI) * LeftStemWidth / 2.0 _
        + Math.Cos(LeftStemTop.angle - Math.PI / 2) * LeftBeamThickness * LeftSpaceHeight / 2.0
        BottomLeftPoint.Y = LeftStemTop.y + Math.Sin(LeftStemTop.angle + Math.PI) * LeftStemWidth / 2.0 _
        + Math.Sin(LeftStemTop.angle - Math.PI / 2) * LeftBeamThickness * LeftSpaceHeight / 2.0

        TopRightPoint.X = RightStemTop.x + Math.Cos(RightStemTop.angle) * RightStemWidth / 2.0 _
        + Math.Cos(RightStemTop.angle + Math.PI / 2) * RightBeamThickness * RightSpaceHeight / 2.0
        TopRightPoint.Y = RightStemTop.y + Math.Sin(RightStemTop.angle) * RightStemWidth / 2.0 _
        + Math.Sin(RightStemTop.angle + Math.PI / 2) * RightBeamThickness * RightSpaceHeight / 2.0

        BottomRightPoint.X = RightStemTop.x + Math.Cos(RightStemTop.angle) * RightStemWidth / 2.0 _
        + Math.Cos(RightStemTop.angle - Math.PI / 2) * RightBeamThickness * RightSpaceHeight / 2.0
        BottomRightPoint.Y = RightStemTop.y + Math.Sin(RightStemTop.angle) * RightStemWidth / 2.0 _
        + Math.Sin(RightStemTop.angle - Math.PI / 2) * RightBeamThickness * RightSpaceHeight / 2.0

    
        'Test: Add those points to the point collections and don't use interpolation


        Dim percentage As Double = 0.0
        Dim TopInterpolatedPoint As bbPoint = Nothing
        Dim BottomInterpolatedPoint As bbPoint = Nothing
        Dim InterpolatedBeamThickness As Double = 0.0
        Dim InterpolatedDisplacement As bbRay = Nothing
        Dim InterpolatedX As Double = 0.0
        Dim InterpolatedY As Double = 0.0
        Dim InterpolatedSpaceHeight As Double = 0.0
        Dim PiDiv2 As Double = Math.PI / 2
        Dim TopAngle As Double = 0.0
        Dim BottomAngle As Double = 0.0
        Dim Amplitude As Double = 0.0

        For i As Integer = 0 To Interpolations
            If i = 0 Then
                BeamTopPoints.Add(TopLeftPoint)
                BeamBottomPoints.Add(BottomLeftPoint)
            ElseIf i = Interpolations Then
                BeamTopPoints.Add(TopRightPoint)
                BeamBottomPoints.Add(BottomRightPoint)
            Else
                percentage = i / Interpolations
                InterpolatedBeamThickness = (RightBeamThickness - LeftBeamThickness) * percentage + LeftBeamThickness
                InterpolatedX = (unitsRightStemX - unitsLeftStemX) * percentage + unitsLeftStemX
                InterpolatedY = (spacesRightStemY - spacesLeftStemY) * percentage + spacesLeftStemY
                Staff.FindPlottingPoint(InterpolatedX, InterpolatedY, InterpolatedDisplacement, InterpolatedSpaceHeight)

                TopAngle = InterpolatedDisplacement.angle + PiDiv2
                BottomAngle = InterpolatedDisplacement.angle - PiDiv2
                Amplitude = InterpolatedBeamThickness * InterpolatedSpaceHeight / 2.0

                TopInterpolatedPoint = New bbPoint
                BottomInterpolatedPoint = New bbPoint

                TopInterpolatedPoint.X = InterpolatedDisplacement.x + Math.Cos(TopAngle) * Amplitude
                TopInterpolatedPoint.Y = InterpolatedDisplacement.y + Math.Sin(TopAngle) * Amplitude
                BottomInterpolatedPoint.X = InterpolatedDisplacement.x + Math.Cos(BottomAngle) * Amplitude
                BottomInterpolatedPoint.Y = InterpolatedDisplacement.y + Math.Sin(BottomAngle) * Amplitude

                BeamTopPoints.Add(TopInterpolatedPoint)
                BeamBottomPoints.Add(BottomInterpolatedPoint)
            End If
        Next i


        'Draw the beam
        Dim s As New System.Text.StringBuilder(1024)
        s.Append(SaveCTM())
        s.Append(PathColor(Color.r, Color.g, Color.b))
        s.Append(PathWidth(0))
        For i As Integer = 1 To BeamTopPoints.Count
            Dim p As bbPoint = BeamTopPoints.Item(i)
            If i = 1 Then
                s.Append(PathStart(p.X, p.Y))
            Else
                s.Append(PathContinue(p.X, p.Y))
            End If
        Next
        For i As Integer = BeamBottomPoints.Count To 1 Step -1
            Dim p As bbPoint = BeamBottomPoints.Item(i)
            s.Append(PathContinue(p.X, p.Y))
        Next

        s.Append(DrawFilledRegion())
        s.Append(RevertCTM())
        Page.AddLineToStream(s.ToString)
    End Sub


    Public Sub DrawSlur( _
                            ByRef Staff As bbStaff, _
 _
                            ByVal unitsLeftX As Double, _
                            ByVal spacesLeftY As Double, _
 _
                            ByVal unitsRightX As Double, _
                            ByVal spacesRightY As Double, _
 _
                            ByVal spacesHeight As Double, _
 _
                            ByVal spacesOuterThickness As Double, _
                            ByVal spacesInnerThickness As Double, _
 _
                            ByVal Interpolations As Integer, _
 _
                            ByVal Color As bbColor)

        'Transform the parabola f(x) = 1-x^2 , across the interval x E [-1, 1] and y E [0, 1]
        Dim SlurTopPoints As New Collection
        Dim SlurMiddlePoints As New Collection
        Dim SlurBottomPoints As New Collection
        Dim Page As bbPDFPage = Nothing
        Staff.GetParent(Page)

        Dim LeftPoint As bbRay = Nothing
        Dim RightPoint As bbRay = Nothing
        Dim LeftSpaceHeight As Double = 0.0
        Dim RightSpaceHeight As Double = 0.0
        Dim SlurAngle As Double

        Staff.FindPlottingPoint(unitsLeftX, spacesLeftY, LeftPoint, LeftSpaceHeight)
        Staff.FindPlottingPoint(unitsRightX, spacesRightY, RightPoint, RightSpaceHeight)
        SlurAngle = New bbRay(New bbPoint(LeftPoint.x, LeftPoint.y), New bbPoint(RightPoint.x, RightPoint.y)).angle + Math.PI / 2

        For i As Integer = 0 To Interpolations
            Dim BaseLineX As Double = LeftPoint.x + (RightPoint.x - LeftPoint.x) * i / Interpolations
            Dim BaseLineY As Double = LeftPoint.y + (RightPoint.y - LeftPoint.y) * i / Interpolations
            Dim BaseLineSpaceHeight As Double = LeftSpaceHeight + (RightSpaceHeight - LeftSpaceHeight) * i / Interpolations
            Dim BaseLineThickness As Double = 0


            If i < Interpolations / 2 Then
                BaseLineThickness = spacesOuterThickness + (spacesInnerThickness - spacesOuterThickness) * i / Interpolations * 2
            Else
                BaseLineThickness = spacesOuterThickness + (spacesInnerThickness - spacesOuterThickness) * (Interpolations - i) / Interpolations * 2
            End If

            BaseLineThickness *= BaseLineSpaceHeight

            Dim ParabolaX As Double = (i / Interpolations * 2 - 1)
            'update this if you change the equation
            Dim ParabolaY As Double = 1 - (Math.Abs(ParabolaX) ^ 3)

            Dim CurveX As Double = BaseLineX + Math.Cos(SlurAngle) * ParabolaY * spacesHeight * BaseLineSpaceHeight
            Dim CurveY As Double = BaseLineY + Math.Sin(SlurAngle) * ParabolaY * spacesHeight * BaseLineSpaceHeight

            Dim TopPointX As Double = CurveX + Math.Cos(SlurAngle) * BaseLineThickness / 2
            Dim TopPointY As Double = CurveY + Math.Sin(SlurAngle) * BaseLineThickness / 2

            Dim BottomPointX As Double = CurveX - Math.Cos(SlurAngle) * BaseLineThickness / 2
            Dim BottomPointY As Double = CurveY - Math.Sin(SlurAngle) * BaseLineThickness / 2

            SlurMiddlePoints.Add(New bbPoint(CurveX, CurveY))
            SlurTopPoints.Add(New bbPoint(TopPointX, TopPointY))
            SlurBottomPoints.Add(New bbPoint(BottomPointX, BottomPointY))
        Next


        'Draw rounded edges
        Dim LeftSlurRay As New bbRay(SlurMiddlePoints(2), SlurMiddlePoints(1))
        Dim RightSlurRay As New bbRay(SlurMiddlePoints(SlurMiddlePoints.Count - 1), SlurMiddlePoints(SlurMiddlePoints.Count))

        Dim h1 As Double = spacesOuterThickness * RightSpaceHeight / 2 '
        Dim lh1 As Double = spacesOuterThickness * LeftSpaceHeight / 2
        Dim theta2 As Double = Math.Abs(SlurAngle + RightSlurRay.angle) '
        Dim ltheta2 As Double = Math.Abs(LeftSlurRay.angle - SlurAngle)
        Dim theta1 As Double = Math.PI / 2 - theta2 '
        Dim ltheta1 As Double = Math.PI / 2 - ltheta2

        Dim lh2 As Double = lh1 * Math.Cos(ltheta1)
        Dim h2 As Double = lh2 '1 * Math.Sin(theta1) '
        Dim REP As bbPoint = SlurBottomPoints(SlurBottomPoints.Count) '
        Dim LEP As bbPoint = SlurBottomPoints(1)
        Dim REPangle As Double = SlurAngle - ltheta2 + Math.PI / 2 '
        Dim LEPangle As Double = SlurAngle + ltheta2 + Math.PI / 2
        Dim REPc As New bbPoint(REP.X + Math.Cos(REPangle) * h2, REP.Y + Math.Sin(REPangle) * h2) '
        Dim LEPc As New bbPoint(LEP.X - Math.Cos(LEPangle) * lh2, LEP.Y - Math.Sin(LEPangle) * lh2)

        Dim RoundedInterpolations As Integer = 20
        Dim addafter As Integer = 1
        For i As Integer = 1 To RoundedInterpolations ' - RoundedInterpolations / 5
            Dim delta As Double = i / RoundedInterpolations * Math.PI
            Dim delta2 As Double = (RoundedInterpolations - i) / RoundedInterpolations * Math.PI

            Dim lp As New bbPoint(LEPc.X, LEPc.Y)
            Dim rp As New bbPoint(REPc.X, REPc.Y)

            Dim la As Double = LEPangle - delta2
            Dim lr As Double = lh2
            Dim ra As Double = REPangle - delta
            Dim rr As Double = h2

            lp.X += Math.Cos(la) * lr
            lp.Y += Math.Sin(la) * lr
            rp.X += Math.Cos(ra) * rr
            rp.Y += Math.Sin(ra) * rr

            SlurBottomPoints.Add(lp, , addafter)
            addafter += 1
            SlurTopPoints.Add(rp)
        Next

        'Draw the slur
        Dim s As New System.Text.StringBuilder(1024)
        s.Append(SaveCTM())
        s.Append(PathColor(Color.r, Color.g, Color.b))
        s.Append(PathWidth(0))
        For i As Integer = 1 To SlurTopPoints.Count
            Dim p As bbPoint = SlurTopPoints.Item(i)
            If i = 1 Then
                s.Append(PathStart(p.X, p.Y))
            Else
                s.Append(PathContinue(p.X, p.Y))
            End If
        Next
        For i As Integer = SlurBottomPoints.Count To 1 Step -1
            Dim p As bbPoint = SlurBottomPoints.Item(i)
            s.Append(PathContinue(p.X, p.Y))
        Next
        s.Append(DrawFilledRegion())
        s.Append(RevertCTM())
        Page.AddLineToStream(s.ToString)
    End Sub
End Module
