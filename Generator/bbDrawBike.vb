imports System.Drawing

Module bbDrawBike
    'Bike frame constants
    Public Const SpaceHeights As Double = 0.02
    Public Const LineThicknesses As Double = SpaceHeights * 0.22
    Public Const BikeFrameRadius As Double = 2.0
    Public Const BikeFrameInnerWidth As Double = 3.1
    Public Const BikeFrameUpperFrontSegmentHeight As Double = 0.3
    Public Const BikeFrameScale As Double = 2.8
    Public BikeFrameUpperBackPosition As New bbPoint(-1.3, BikeFrameRadius * 2.0 + 1.0)
    Public BikeFramePedalOrigin As New bbPoint(-0.1, -0.7 + BikeFrameRadius)
    Public BikeFrameOrigin As New bbPoint(32.0 / 2.0, 1.6)
    Public Const GroundWidth As Double = 4.0
    Public Const GroundYDisplacement As Double = SpaceHeights * -10.5
    Public Const GroundLineExtension As Double = 5.5
    Public Const LeftGearsArcWidth As Double = 0.5


    'Colors
    Public FrameColor As New bbColor(0.9, 0.9, 0.94) 'red appears 0.03 more
    Public WheelColor As New bbColor(0.2, 0.3, 0.4)
    Public BikeGuardColor As New bbColor(0.2, 0.3, 0.4)
    Public LeftOuterGearColor As New bbColor(0.2, 0.3, 0.4)
    Public LeftInnerGearColor As New bbColor(0.2, 0.3, 0.4)
    Public RightGearColor As New bbColor(0.2, 0.3, 0.4)

    Public GroundLineColor As New bbColor(0.2, 0.3, 0.4)
    Public GroundColors As New bbColor(0.2, 0.3, 0.4)

    Dim LeftCenter As bbStaff
    Dim RightCenter As bbStaff
    Dim CenterPoints As Integer = 200

    Dim LeftBikeWheel As bbStaff
    Dim LeftBikeWheelPoints As Integer = 200
    Dim RightBikeWheel As bbStaff
    Dim RightBikeWheelPoints As Integer = 200


    Dim LeftBikeGuard As bbStaff
    Dim LeftBikeGuardPoints As Integer = 200

    Dim RightBikeGuard As bbStaff
    Dim RightBikeGuardPoints As Integer = LeftBikeGuardPoints / 2

    Dim BikeGuardRadius As Double = BikeFrameRadius + 0.35 'was BFR + 0.3

    Dim LeftBikeGear As bbStaff
    Dim LeftBikeGearPoints As Integer = 400
    Dim LeftBikeInnerGear As bbStaff
    Dim LeftBikeInnerGearPoints As Integer = 400

    Dim RightBikeGear As bbStaff
    Dim RightBikeGearPoints As Integer = 359
    Dim RightBikeInnerGear As bbStaff
    Dim RightBikeInnerGearPoints As Integer = 200

    Dim GroundLeft As bbStaff
    Dim GroundRight As bbStaff
    Dim GroundLine As bbStaff
    Public Const GroundPoints As Integer = 10

    Public Const LeftBikeGearRadius As Double = 0.4
    Public Const LeftBikeInnerGearRadius As Double = 0.25
    Public Const RightBikeGearRadius As Double = 0.6
    Public Const RightBikeInnerGearRadius As Double = 0.45

    Public Const LeftSpokeRadius As Double = 0.4 + SpaceHeights * 2.0
    Public Const RightSpokeRadius As Double = 0.1
    Public Const SpokeSlant As Double = Math.PI / 16
    Public Const SpokesPerWheel As Integer = 40
    Public Const SpokeThickness As Double = 0.003

    'Handlebars and seat
    Public Const HandleBarsHeight As Double = 0.45
    Public Const HandleBarsExtension As Double = 0.45
    Public Const HandleBarsRadius As Double = 0.38
    Public Const HandleBarsCirclePercentage As Double = 0.55
    Public Const HandleBarsCirclePoints As Integer = 200

    Public Const SeatHeight As Double = 0.35
    Public Const SeatWidth As Double = 0.75
    Public Const SeatTipAngle As Double = Math.PI - Math.PI / 3
    Public Const SeatTipLength As Double = 0.2
    Public Const SeatBackWidth = 0.2

    Public Sub DrawBikeFrame(ByRef Page As bbPDFPage)
        StaffLineTime.StartClock()
        Dim PointOrigin As New bbPoint(0, 0)
        Dim Point1 As New bbPoint(-BikeFrameInnerWidth, BikeFrameRadius)
        Dim Point2 As New bbPoint(BikeFrameInnerWidth, BikeFrameRadius)
        Dim Point3 As New bbPoint(BikeFramePedalOrigin.X, BikeFramePedalOrigin.Y)
        Dim Point4 As New bbPoint(BikeFrameUpperBackPosition.X, BikeFrameUpperBackPosition.Y)
        Dim slope As Double = (Point4.Y - Point3.Y) / (Point4.X - Point3.X)
        Dim x = (Point4.Y - Point2.Y) / slope + Point2.X
        Dim Point5 As New bbPoint(x, Point4.Y)
        Dim p2_5_i As Double = 0.91
        Dim Point6 As New bbPoint((Point5.X - Point2.X) * p2_5_i + Point2.X, (Point5.Y - Point2.Y) * p2_5_i + Point2.Y)


        'Handle bars (pt. 7 to pt. 9)
        Dim Ray5to6 As New bbRay(Point5, Point6)
        Dim UpwardsAngle As Double = Ray5to6.angle + Math.PI
        Dim Point7 As New bbPoint(Point5.X + Math.Cos(UpwardsAngle) * HandleBarsHeight, Point5.Y + Math.Sin(UpwardsAngle) * HandleBarsHeight)
        Dim PerpendicularAngle = UpwardsAngle - Math.PI / 2
        Dim Point8 As New bbPoint(Point7.X + Math.Cos(PerpendicularAngle) * HandleBarsExtension, Point7.Y + Math.Sin(PerpendicularAngle) * HandleBarsExtension)
        Dim PerpendicularAngle2 = PerpendicularAngle - Math.PI / 2
        Dim Point9 As New bbPoint(Point8.X + Math.Cos(PerpendicularAngle2) * HandleBarsRadius, Point8.Y + Math.Sin(PerpendicularAngle2) * HandleBarsRadius)


        Dim Ray3To4 As New bbRay(Point3, Point4)
        Dim Point10 As New bbPoint(Point4.X + Math.Cos(Ray3To4.angle) * SeatHeight, Point4.Y + Math.Sin(Ray3To4.angle) * SeatHeight)
        Dim Point10A As New bbPoint(Point10.X - SeatBackWidth, Point10.Y)
        Dim Point11 As New bbPoint(Point10.X + SeatWidth, Point10.Y)
        Dim Point12 As New bbPoint(Point10A.X + Math.Cos(SeatTipAngle) * SeatTipLength, Point10A.Y + Math.Sin(SeatTipAngle) * SeatTipLength)





        Page.AddLineToStream(CTMTranslate(BikeFrameOrigin.X, BikeFrameOrigin.Y))
        Page.AddLineToStream(CTMScale(BikeFrameScale, BikeFrameScale))

        'Page.AddLineToStream(DrawPoint(PointOrigin, 0.1))

        Page.AddLineToStream(SaveCTM())

        Page.AddLineToStream(PathCap(True))

        Page.AddLineToStream(DrawLine(Point1, Point4, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point4, Point5, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point5, Point2, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point4, Point3, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point3, Point6, 0.1, FrameColor))

        'Handlebars
        Page.AddLineToStream(DrawLine(Point5, Point7, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point7, Point8, 0.1, FrameColor))

        'Seat
        Page.AddLineToStream(DrawLine(Point4, Point10, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point10A, Point11, 0.1, FrameColor))
        Page.AddLineToStream(DrawLine(Point10A, Point12, 0.1, FrameColor))

        Dim DeltaAngle As Double = UpwardsAngle

        Dim Hx As Double = Point9.X + Math.Cos(DeltaAngle) * HandleBarsRadius
        Dim Hy As Double = Point9.Y + Math.Sin(DeltaAngle) * HandleBarsRadius

        Dim HxOld As Double
        Dim HyOld As Double

        For i As Integer = 1 To HandleBarsCirclePoints
            HxOld = Hx
            HyOld = Hy
            DeltaAngle = UpwardsAngle - Math.PI * 2 * HandleBarsCirclePercentage * i / HandleBarsCirclePoints
            Hx = Point9.X + Math.Cos(DeltaAngle) * HandleBarsRadius
            Hy = Point9.Y + Math.Sin(DeltaAngle) * HandleBarsRadius

            Dim PointL As New bbPoint(HxOld, HyOld)
            Dim PointR As New bbPoint(Hx, Hy)
            Page.AddLineToStream(DrawLine(PointL, PointR, 0.1, FrameColor))
        Next

        Page.AddLineToStream(RevertCTM())

        'Spokes
        For i As Integer = 0 To SpokesPerWheel - 1
            Dim OuterAngle As Double = i / SpokesPerWheel * Math.PI * 2.0
            Dim InnerAngle As Double = 0.0

            If i Mod 2 = 0 Then
                InnerAngle = OuterAngle + SpokeSlant
            Else
                InnerAngle = OuterAngle - SpokeSlant
            End If

            Dim InnerPointL As New bbPoint(Point1.X + Math.Cos(InnerAngle) * RightSpokeRadius, Point1.Y + Math.Sin(InnerAngle) * RightSpokeRadius)
            Dim OuterPointL As New bbPoint(Point1.X + Math.Cos(OuterAngle) * (BikeFrameRadius - SpaceHeights * 2.0), Point1.Y + Math.Sin(OuterAngle) * (BikeFrameRadius - SpaceHeights * 2.0))

            Dim InnerPointR As New bbPoint(Point2.X + Math.Cos(InnerAngle) * RightSpokeRadius, Point2.Y + Math.Sin(InnerAngle) * RightSpokeRadius)
            Dim OuterPointR As New bbPoint(Point2.X + Math.Cos(OuterAngle) * (BikeFrameRadius - SpaceHeights * 2.0), Point2.Y + Math.Sin(OuterAngle) * (BikeFrameRadius - SpaceHeights * 2.0))

            Page.AddLineToStream(DrawLine(InnerPointL, OuterPointL, SpokeThickness))
            Page.AddLineToStream(DrawLine(InnerPointR, OuterPointR, SpokeThickness))
        Next
        Page.AddLineToStream(DrawPoint(Point1.X, Point1.Y, LeftSpokeRadius * 2.0, New bbColor(1.0, 1.0, 1.0)))


        'Left bike wheel
        LeftBikeWheel = New bbStaff(Page)
        For i As Integer = -2 To 2
            LeftBikeWheel.StaffProperty(0).StaffLine(i).Draw.Value = True
            LeftBikeWheel.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            LeftBikeWheel.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = WheelColor
        Next
        LeftBikeWheel.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To LeftBikeWheelPoints
            LeftBikeWheel.Point(i).Angle = -i / LeftBikeWheelPoints * Math.PI * 2.0 + Math.PI / 2.0
            LeftBikeWheel.Point(i).Radius = BikeFrameRadius
            LeftBikeWheel.Point(i).X += Point1.X
            LeftBikeWheel.Point(i).Y += Point1.Y
        Next
        LeftBikeWheel.DrawStaff()

        'Right bike wheel
        RightBikeWheel = New bbStaff(Page)
        For i As Integer = -2 To 2
            RightBikeWheel.StaffProperty(0).StaffLine(i).Draw.Value = True
            RightBikeWheel.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            RightBikeWheel.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = WheelColor
        Next
        RightBikeWheel.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To RightBikeWheelPoints
            RightBikeWheel.Point(i).Angle = -i / RightBikeWheelPoints * Math.PI * 2.0 + Math.PI / 2.0
            RightBikeWheel.Point(i).Radius = BikeFrameRadius
            RightBikeWheel.Point(i).X += Point2.X
            RightBikeWheel.Point(i).Y += Point2.Y
        Next
        RightBikeWheel.DrawStaff()

        'Right bike guard
        RightBikeGuard = New bbStaff(Page)
        For i As Integer = -2 To 2
            RightBikeGuard.StaffProperty(0).StaffLine(i).Draw.Value = False
            RightBikeGuard.StaffProperty(Fix(RightBikeGuardPoints / 12)).StaffLine(i).Draw.Value = True

            RightBikeGuard.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            RightBikeGuard.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = BikeGuardColor
        Next
        RightBikeGuard.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To RightBikeGuardPoints
            RightBikeGuard.Point(i).Angle = -i / RightBikeGuardPoints * Math.PI / 2.0 + Math.PI
            RightBikeGuard.Point(i).Radius = BikeGuardRadius
            RightBikeGuard.Point(i).X += Point2.X
            RightBikeGuard.Point(i).Y += Point2.Y
        Next
        RightBikeGuard.DrawStaff()


        'Left bike guard
        LeftBikeGuard = New bbStaff(Page)
        Dim lbg_halfway_point As Integer = LeftBikeGuardPoints / 2
        For i As Integer = -2 To 2
            LeftBikeGuard.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            LeftBikeGuard.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = BikeGuardColor
            If i <> 0 Then
                LeftBikeGuard.StaffProperty(0).StaffLine(i).Draw.Value = False
                LeftBikeGuard.StaffProperty(1).StaffLine(i).Draw.Value = False
            Else
                LeftBikeGuard.StaffProperty(0).StaffLine(i).DashOffLength.Value = 0.009
                LeftBikeGuard.StaffProperty(0).StaffLine(i).DashOnLength.Value = 0.0
                LeftBikeGuard.StaffProperty(0).StaffLine(i).DashOffset.Value = 0.0
                LeftBikeGuard.StaffProperty(0).StaffLine(i).DrawDashedLine.Value = True
            End If

            LeftBikeGuard.StaffProperty(lbg_halfway_point).StaffLine(i).Draw.Value = True
            LeftBikeGuard.StaffProperty(lbg_halfway_point).StaffLine(i).DrawDashedLine.Value = False
        Next
        LeftBikeGuard.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To LeftBikeGuardPoints
            LeftBikeGuard.Point(i).Angle = -i / LeftBikeGuardPoints * Math.PI + Math.PI
            LeftBikeGuard.Point(i).Radius = BikeGuardRadius
            LeftBikeGuard.Point(i).X += Point1.X
            LeftBikeGuard.Point(i).Y += Point1.Y
        Next
        LeftBikeGuard.DrawStaff()




        'Left gears
        LeftBikeGear = New bbStaff(Page)
        LeftBikeGear.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        'c=2*pi*r
        'p = w/(2*pi*r)
        Dim dashed_line_start As Integer = LeftBikeGearPoints * (LeftGearsArcWidth / (2 * Math.PI * LeftBikeGearRadius))
        Dim dashed_inner_line_start As Integer = LeftBikeGearPoints * (LeftGearsArcWidth / (2 * Math.PI * LeftBikeInnerGearRadius))

        For i As Integer = -2 To 2
            LeftBikeGear.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            LeftBikeGear.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = LeftOuterGearColor
            If i <> 0 Then
                LeftBikeGear.StaffProperty(dashed_line_start).StaffLine(i).Draw.Value = False
            Else
                LeftBikeGear.StaffProperty(dashed_line_start).StaffLine(i).DashOffLength.Value = 0.009
                LeftBikeGear.StaffProperty(dashed_line_start).StaffLine(i).DashOnLength.Value = 0.0
                LeftBikeGear.StaffProperty(dashed_line_start).StaffLine(i).DashOffset.Value = 0.0
                LeftBikeGear.StaffProperty(dashed_line_start).StaffLine(i).DrawDashedLine.Value = True
            End If
        Next

        For i As Integer = 0 To LeftBikeGearPoints
            With LeftBikeGear.Point(i)
                .Angle = -i / LeftBikeGearPoints * Math.PI * 2.0 - Math.PI * 1.5
                .Radius = LeftBikeGearRadius
                .X += Point1.X
                .Y += Point1.Y
            End With
        Next
        LeftBikeGear.DrawStaff()

        LeftBikeInnerGear = New bbStaff(Page)

        LeftBikeInnerGear.StaffProperty(0).SpaceHeight.Value = SpaceHeights
        For i As Integer = -2 To 2
            LeftBikeInnerGear.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            LeftBikeInnerGear.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = LeftInnerGearColor
            If i <> 0 Then
                LeftBikeInnerGear.StaffProperty(dashed_inner_line_start).StaffLine(i).Draw.Value = False
            Else
                LeftBikeInnerGear.StaffProperty(dashed_inner_line_start).StaffLine(i).DashOffLength.Value = 0.009
                LeftBikeInnerGear.StaffProperty(dashed_inner_line_start).StaffLine(i).DashOnLength.Value = 0.0
                LeftBikeInnerGear.StaffProperty(dashed_inner_line_start).StaffLine(i).DashOffset.Value = 0.0
                LeftBikeInnerGear.StaffProperty(dashed_inner_line_start).StaffLine(i).DrawDashedLine.Value = True
            End If
        Next


        For i As Integer = 0 To LeftBikeInnerGearPoints
            With LeftBikeInnerGear.Point(i)
                .Angle = -i / LeftBikeInnerGearPoints * Math.PI * 2.0 - Math.PI * 1.65
                .Radius = LeftBikeInnerGearRadius
                .X += Point1.X
                .Y += Point1.Y
            End With
        Next
        LeftBikeInnerGear.DrawStaff()


        'Right gears
        RightBikeGear = New bbStaff(Page)
        For i As Integer = -2 To 2
            RightBikeGear.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            RightBikeGear.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = RightGearColor
        Next
        RightBikeGear.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To RightBikeGearPoints
            With RightBikeGear.Point(i)
                .Angle = -i / RightBikeGearPoints * Math.PI * 2.0 + Math.PI / 2
                .Radius = RightBikeGearRadius
                .X += Point3.X
                .Y += Point3.Y
            End With
        Next
        RightBikeGear.DrawStaff()

        RightBikeInnerGear = New bbStaff(Page)
        Dim dashed_line_right_gear_stop As Integer = RightBikeInnerGearPoints * 0.42
        For i As Integer = -2 To 2
            RightBikeInnerGear.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            RightBikeInnerGear.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = RightGearColor
            If i <> 0 Then
                RightBikeInnerGear.StaffProperty(dashed_line_right_gear_stop).StaffLine(i).Draw.Value = False
            Else
                RightBikeInnerGear.StaffProperty(dashed_line_right_gear_stop).StaffLine(i).DashOffLength.Value = 0.009
                RightBikeInnerGear.StaffProperty(dashed_line_right_gear_stop).StaffLine(i).DashOnLength.Value = 0.0
                RightBikeInnerGear.StaffProperty(dashed_line_right_gear_stop).StaffLine(i).DashOffset.Value = 0.0
                RightBikeInnerGear.StaffProperty(dashed_line_right_gear_stop).StaffLine(i).DrawDashedLine.Value = True
            End If
        Next

        RightBikeInnerGear.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To RightBikeInnerGearPoints
            With RightBikeInnerGear.Point(i)
                .Angle = (-i / RightBikeInnerGearPoints + 0.25 - 2 / 3) * Math.PI * 2.0
                .Radius = RightBikeInnerGearRadius
                .X += Point3.X
                .Y += Point3.Y
            End With
        Next
        RightBikeInnerGear.DrawStaff()


        'Ground
        GroundLeft = New bbStaff(Page)
        GroundRight = New bbStaff(Page)
        GroundLine = New bbStaff(Page)

        GroundLeft.StaffProperty(0).SpaceHeight.Value = SpaceHeights
        GroundRight.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        GroundLine.StaffProperty(0).StaffLine(0).Draw.Value = True
        GroundLine.StaffProperty(0).StaffLine(0).LineThickness.Value = LineThicknesses
        GroundLine.StaffProperty(0).StaffLine(0).LineStrokeColor.Value = GroundLineColor

        For i As Integer = -2 To 2
            GroundLeft.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            GroundLeft.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = GroundColors
            GroundRight.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            GroundRight.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = GroundColors
        Next

        For i As Integer = -2 To 2
            Dim p As Integer = (0.0) * GroundPoints
            Dim value As Boolean
            If i = 0 Then
                value = True
            Else
                value = False
            End If
            GroundLeft.StaffProperty(p).StaffLine(i).Draw.Value = True
            GroundRight.StaffProperty(p).StaffLine(i).Draw.Value = True
        Next

        For i As Integer = 0 To GroundPoints
            GroundLeft.Point(i).X = Point1.X - GroundWidth / 2.0 + (i / GroundPoints * GroundWidth)
            GroundLeft.Point(i).Y = GroundYDisplacement
            GroundRight.Point(i).X = Point2.X - GroundWidth / 2.0 + (i / GroundPoints * GroundWidth)
            GroundRight.Point(i).Y = GroundYDisplacement
        Next

        For i As Integer = 0 To GroundPoints
            GroundLine.Point(i).X = (Point2.X + Point1.X) / 2.0 - GroundLineExtension + (i / GroundPoints * GroundLineExtension * 2.0)
            GroundLine.Point(i).Y = GroundYDisplacement ' + SpaceHeights * 2.0
        Next

        GroundLine.DrawStaff()
        GroundLeft.DrawStaff()
        GroundRight.DrawStaff()

        'Center circles
        LeftCenter = New bbStaff(Page)
        RightCenter = New bbStaff(Page)
        For i As Integer = -2 To 2
            If i = 0 Then
                LeftCenter.StaffProperty(0).StaffLine(i).Draw.Value = True
                RightCenter.StaffProperty(0).StaffLine(i).Draw.Value = True
            Else
                LeftCenter.StaffProperty(0).StaffLine(i).Draw.Value = False
                RightCenter.StaffProperty(0).StaffLine(i).Draw.Value = False
            End If
            LeftCenter.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            LeftCenter.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = New bbColor(Color.Black)

            RightCenter.StaffProperty(0).StaffLine(i).LineThickness.Value = LineThicknesses
            RightCenter.StaffProperty(0).StaffLine(i).LineStrokeColor.Value = New bbColor(Color.Black)
        Next
        LeftCenter.StaffProperty(0).SpaceHeight.Value = SpaceHeights
        RightCenter.StaffProperty(0).SpaceHeight.Value = SpaceHeights

        For i As Integer = 0 To CenterPoints
            LeftCenter.Point(i).Angle = -i / CenterPoints * Math.PI * 2.0 + Math.PI / 2.0
            LeftCenter.Point(i).Radius = RightSpokeRadius
            LeftCenter.Point(i).X += Point1.X
            LeftCenter.Point(i).Y += Point1.Y

            RightCenter.Point(i).Angle = -i / CenterPoints * Math.PI * 2.0 + Math.PI / 2.0
            RightCenter.Point(i).Radius = RightSpokeRadius
            RightCenter.Point(i).X += Point2.X
            RightCenter.Point(i).Y += Point2.Y
        Next
        LeftCenter.DrawStaff()
        RightCenter.DrawStaff()
        StaffLineTime.StopClock()

        'Test text***********
        Dim EspressText As New bbText(GillSansItalic, 0.045)
        'EspressText.DrawHorizontalToStaff(LeftBikeWheel, "At a pedaling tempo", 0.0, 12.0, bbText.bbJustification.bbJustifyLeft)
        Dim BoxText As New bbText(GillSansRegular, 0.045)
        Dim SymbolText As New bbText(NotationFont, 0.045)

        BoxText.DrawBoxed(LeftBikeWheel, "A", 0, 28 - 3, 0.002, 0.2)
        BoxText.DrawBoxed(LeftBikeWheel, "C", 0, 22 - 3, 0.002, 0.2)
        BoxText.DrawBoxed(LeftBikeWheel, "F", 0, 16 - 3, 0.002, 0.2)
        BoxText.DrawBoxed(LeftBikeWheel, "H", 0, 10 - 3, 0.002, 0.2)

        BoxText.DrawBoxed(LeftBikeWheel, "B", LeftBikeWheel.StaffLength() / 4, 11, 0.002, 0.2, 0.08)

        BoxText.DrawBoxed(LeftBikeWheel, "I", LeftBikeWheel.StaffLength() / 2 + 0.05, 6, 0.002, 0.2)
        BoxText.DrawBoxed(LeftBikeWheel, "D", LeftBikeWheel.StaffLength() / 2 + 0.05, 12, 0.002, 0.2)

        BoxText.DrawBoxed(LeftBikeWheel, "G", RightBikeWheel.StaffLength() * 3 / 4, 11, 0.002, 0.2, 0.08)

        BoxText.DrawBoxed(GroundLeft, "L", 0.08, 8.7, 0.002, 0.2)
        BoxText.DrawBoxed(LeftBikeGuard, "J", LeftBikeGuard.StaffLength() / 2.0 + 0.1, 13, 0.002, 0.3, 0.05)

        BoxText.DrawBoxed(LeftBikeGear, "E", 0.05, 9, 0.002, 0.2, 0.03)
        BoxText.DrawBoxed(LeftBikeInnerGear, "K", 0.038, 5.5, 0.002, 0.2, 0.08)
        EspressText.DrawHorizontalToStaff(LeftBikeGear, "ost.", 0.079, 9, bbText.bbJustification.bbJustifyLeft)
        EspressText.DrawHorizontalToStaff(LeftBikeInnerGear, "ost.", 0.068, 5.5, bbText.bbJustification.bbJustifyLeft)
        EspressText.DrawHorizontalToStaff(LeftBikeGuard, "Molto accelerando!!!", LeftBikeGuard.StaffLength() / 2 + 0.13, 13, bbText.bbJustification.bbJustifyLeft)


        BoxText.DrawBoxed(RightBikeWheel, "A", 0.0, 28 - 3, 0.002, 0.2)
        BoxText.DrawBoxed(RightBikeWheel, "C", 0.0, 22 - 3, 0.002, 0.2)
        BoxText.DrawBoxed(RightBikeWheel, "F", 0.0, 16 - 3, 0.002, 0.2)
        BoxText.DrawBoxed(RightBikeWheel, "H", 0.0, 10 - 3, 0.002, 0.2)


        BoxText.DrawBoxed(RightBikeWheel, "B", RightBikeWheel.StaffLength() / 4, 7, 0.002, 0.2, 0.05)

        BoxText.DrawBoxed(RightBikeWheel, "I", RightBikeWheel.StaffLength() / 2 + 0.05, 6, 0.002, 0.2)

        EspressText.DrawHorizontalToStaff(RightBikeWheel, "From", RightBikeWheel.StaffLength() * 26 / 40, 27, bbText.bbJustification.bbJustifyCenter)
        EspressText.DrawHorizontalToStaff(LeftBikeWheel, "From", LeftBikeWheel.StaffLength() * 26 / 40, 27, bbText.bbJustification.bbJustifyCenter)
        EspressText.DrawHorizontalToStaff(RightBikeWheel, "go to", RightBikeWheel.StaffLength() * 26 / 40, 15, bbText.bbJustification.bbJustifyCenter)
        EspressText.DrawHorizontalToStaff(LeftBikeWheel, "go to", LeftBikeWheel.StaffLength() * 26 / 40, 15, bbText.bbJustification.bbJustifyCenter)
        BoxText.DrawBoxed(LeftBikeWheel, "I", LeftBikeWheel.StaffLength() * 26 / 40 + 0.0, 21, 0.002, 0.2, 0.08)
        BoxText.DrawBoxed(RightBikeWheel, "I", RightBikeWheel.StaffLength() * 26 / 40 + 0.0, 21, 0.002, 0.2, 0.08)
        BoxText.DrawBoxed(LeftBikeWheel, "J", LeftBikeWheel.StaffLength() * 26 / 40 + 0.0, 9, 0.002, 0.3, 0.05)
        BoxText.DrawBoxed(RightBikeWheel, "J", RightBikeWheel.StaffLength() * 26 / 40 + 0.0, 9, 0.002, 0.3, 0.05)

        BoxText.DrawBoxed(RightBikeWheel, "D", RightBikeWheel.StaffLength() / 2 + 0.05, 12, 0.002, 0.2)

        BoxText.DrawBoxed(RightBikeWheel, "G", RightBikeWheel.StaffLength() * 3 / 4, 7, 0.002, 0.2, 0.05)


        BoxText.DrawBoxed(GroundRight, "L", 0.08, 8.7, 0.002, 0.2)
        BoxText.DrawBoxed(RightBikeGuard, "J", RightBikeGuard.StaffLength() - 0.05, 18, 0.002, 0.3, 0.05)
        EspressText.DrawHorizontalToStaff(RightBikeGuard, "Molto accelerando!!!", RightBikeGuard.StaffLength() - 0.08, 18, bbText.bbJustification.bbJustifyRight)

        BoxText.DrawBoxed(RightBikeGear, "K", 0.04, 19, 0.002, 0.2, 0.08)
        BoxText.DrawBoxed(RightBikeGear, "E", 0.038, 25, 0.002, 0.2, 0.03)

        Dim pp As bbPDFPage = Nothing
        RightBikeGear.GetParent(pp)
        Dim O As bbRay = Nothing
        Dim gearspaceheight As Double
        RightBikeGear.FindPlottingPoint(0.04, -10, O, gearspaceheight, 0)
        pp.AddLineToStream(DrawString(NotationFont, gearspaceheight, O, "y"))

        RightBikeWheel.FindPlottingPoint(0.04, -16, O, gearspaceheight, 0)
        pp.AddLineToStream(DrawString(NotationFont, gearspaceheight, O, "y"))

        RightBikeWheel.FindPlottingPoint(0.04 + RightBikeWheel.StaffLength() / 40.0, -16, O, gearspaceheight, 0)
        pp.AddLineToStream(DrawString(NotationFont, gearspaceheight, O, "y"))


        'SymbolText.DrawHorizontalToStaff(RightBikeGear, "y", 0.04, -19, bbText.bbJustification.bbJustifyCenter)

        ' SEGMENT KERNED
        Dim ExtraKern As Double = 0.2
        EspressText.DrawHorizontalToStaff(RightBikeGear, "Two and two-thirds times", 0.078, 25, bbText.bbJustification.bbJustifyLeft, 0.8) '
        EspressText.DrawHorizontalToStaff(RightBikeGear, "Just once", 0.081, 19, bbText.bbJustification.bbJustifyLeft, 0.3) '
        EspressText.DrawHorizontalToStaff(RightBikeWheel, "sim.", RightBikeWheel.StaffLength() / 40 * 2 + 0.05, -16, bbText.bbJustification.bbJustifyLeft, 0)
        EspressText.DrawHorizontalToStaff(RightBikeGear, "sempre", 0.15, -9, bbText.bbJustification.bbJustifyLeft, 0)
        EspressText.DrawHorizontalToStaff(RightBikeWheel, " Blissful,  melodic;  at a pedaling tempo...", 0.03, 28 - 3, bbText.bbJustification.bbJustifyLeft, ExtraKern)

        Dim BoxText2 As New bbText(FontinSmallCaps, 0.07)
        Dim BoxText3 As New bbText(FontinSmallCaps, 0.04)
        BoxText2.DrawHorizontalToStaff(GroundLeft, "B.H.", GroundLeft.StaffLength() - 0.005, 7, bbText.bbJustification.bbJustifyRight)
        BoxText2.DrawHorizontalToStaff(GroundRight, "M.H.", GroundRight.StaffLength() - 0.005, 7, bbText.bbJustification.bbJustifyRight)

        '********************

        NotesTime.StartClock()

        'Draw notes
        DrawLeftBikeWheel(LeftBikeWheel)
        DrawRightBikeWheel(RightBikeWheel)
        DrawLeftBikeGear(LeftBikeGear)
        DrawLeftBikeInnerGear(LeftBikeInnerGear)
        DrawRightBikeGear(RightBikeGear)
        DrawRightBikeInnerGear(RightBikeInnerGear)
        DrawLeftBikeGuard(LeftBikeGuard)
        DrawRightBikeGuard(RightBikeGuard)
        DrawLeftGroundNotes(GroundLeft)
        DrawRightGroundNotes(GroundRight)

        'Draw slurs
        DrawSlur(LeftBikeWheel, 1.3, 10, 1.9, 11, 3, 0.3, 0.6, 100, New bbColor(0, 0, 0))

        NotesTime.StopClock()
    End Sub
End Module
