Public Class bbStaff
    Private StaffPathPoints As New bbSmartCollection 'bbPoint collection
    Private StaffProperties As New bbSmartCollection 'bbStaffProperty collection

    Private Parent As bbPDFPage

    Private PointsPrecomputed As Boolean = False
    Private PrecomputedInfo() As bbStaffPointInfo
    Private VirtualPoints As Integer
    Private ActualPoints As Integer
    Private Log2ActualPoints As Integer

    Private HighestLinePrecomputed As Boolean = False
    Private HighestLine As Integer
    Private LowestLinePrecomputed As Boolean = False
    Private LowestLine As Integer

    Public Sub New(ByRef ParentPage As bbPDFPage)
        Parent = ParentPage
    End Sub


    Public Property Point(ByVal Index As Integer) As bbPoint
        Get
            Dim x As bbPoint = StaffPathPoints(Index)
            If IsNothing(x) Then
                'We have to return something, and we have to make 
                'sure it gets added to the collection
                x = New bbPoint
                StaffPathPoints(Index) = x
            End If
            Return x
        End Get
        Set(ByVal value As bbPoint)
            StaffPathPoints(Index) = value
        End Set
    End Property

    Public Sub GetParent(ByRef Page As bbPDFPage)
        Page = Parent
    End Sub

    Public Function StaffLength() As Double
        If Not PointsPrecomputed Then PrecomputePoints()
        Dim TotalDistance As Double = 0.0
        TotalDistance = PrecomputedInfo(VirtualPoints - 1).RightDistance
        Return TotalDistance
    End Function

    Private Structure bbStaffPointInfo
        Public LeftPoint As Integer
        Public RightPoint As Integer
        Public LeftStaffHeight As bbCascadingProperty
        Public RightStaffHeight As bbCascadingProperty
        Public LeftDistance As Double
        Public RightDistance As Double
    End Structure

    Private Sub PrecomputePoints()
        VirtualPoints = StaffPathPoints.UBound() - StaffPathPoints.LBound()
        Dim x As Double = VirtualPoints
        Dim k As Integer = 0
        Do While x > 1.0
            x /= 2.0
            k += 1
        Loop
        Log2ActualPoints = k
        ActualPoints = 2 ^ Log2ActualPoints
        ReDim PrecomputedInfo(ActualPoints - 1)
        With PrecomputedInfo(0)
            .LeftStaffHeight = New bbCascadingProperty
            .RightStaffHeight = New bbCascadingProperty
            .LeftStaffHeight.Value = 0.0
            .RightStaffHeight.Value = 0.0
        End With

        Dim LastRDistance As Double
        For i As Integer = 0 To StaffPathPoints.UBound() - StaffPathPoints.LBound() - 1
            With PrecomputedInfo(i)
                .LeftPoint = i + StaffPathPoints.LBound()
                .RightPoint = .LeftPoint + 1
                If i <> 0 Then
                    .LeftStaffHeight = PrecomputedInfo(i - 1).RightStaffHeight
                    .RightStaffHeight = .LeftStaffHeight + StaffProperty(.RightPoint).SpaceHeight
                    .LeftDistance = PrecomputedInfo(i - 1).RightDistance
                Else
                    .LeftStaffHeight = StaffProperty(.LeftPoint).SpaceHeight
                    .RightStaffHeight = StaffProperty(.LeftPoint).SpaceHeight
                    .RightStaffHeight += StaffProperty(.RightPoint).SpaceHeight
                    .LeftDistance = 0.0
                End If
                .RightDistance = .LeftDistance + Math.Sqrt((Point(.RightPoint).X - Point(.LeftPoint).X) ^ 2 + (Point(.RightPoint).Y - Point(.LeftPoint).Y) ^ 2)
                LastRDistance = .RightDistance
            End With
        Next
        For i As Integer = StaffPathPoints.UBound() - StaffPathPoints.LBound() To ActualPoints - 1
            PrecomputedInfo(i).LeftDistance = PrecomputedInfo(i - 1).RightDistance 'arbitrary delta, just enough to make the binary search not get confused
            PrecomputedInfo(i).RightDistance = PrecomputedInfo(i).LeftDistance + 1.0 'arbitrary delta, just enough to make the binary search not get confused
        Next
        PointsPrecomputed = True
    End Sub

    Private Function FindPoint(ByVal unitsX As Double) As bbStaffPointInfo
        If Not PointsPrecomputed Then PrecomputePoints()
        'For i As Integer = 0 To ActualPoints - 1
        ' If unitsX < PrecomputedInfo(i).RightDistance Then
        'Return PrecomputedInfo(i)
        'End If
        'Next
        'Return PrecomputedInfo(0)
        'Exit Function
        Dim left As Integer = 0
        Dim right As Integer = ActualPoints - 1
        Dim mean As Integer
        For i As Integer = 0 To Log2ActualPoints
            mean = (right - left + 1) / 2 + left - 1
            With PrecomputedInfo(mean)
                If unitsX < .RightDistance Then
                    If .LeftDistance <= unitsX Then
                        'Found the spot
                        Return PrecomputedInfo(mean)
                    Else
                        'Overshot to the right, so go left
                        left = left
                        right = mean
                    End If
                Else
                    'Overshot to the left, so go right
                    left = mean + 1
                    right = right
                End If
            End With
        Next
        Return PrecomputedInfo(0)
    End Function

    Public Sub FindPlottingPoint(ByVal unitsX As Double, ByVal spacesY As Double, ByRef unitsOrientation As bbRay, ByRef unitsSpaceHeight As Double, ByRef LeftStaffPoint As Integer)
        FindPlottingPointTime.StartClock()
        'The first thing to do is find between which two points the horizontal component exists
        Dim info As bbStaffPointInfo = FindPoint(unitsX)

        Dim LeftPoint As Integer = info.LeftPoint
        Dim RightPoint As Integer = info.RightPoint
        Dim LeftDistance As Double = info.LeftDistance
        Dim RightDistance As Double = info.RightDistance
        Dim LeftStaffHeight As Double = info.LeftStaffHeight.Value
        Dim RightStaffHeight As Double = info.RightStaffHeight.Value
        LeftStaffPoint = LeftPoint
        Dim InterpolatedPoint As New bbPoint
        Dim InterpolatedPercentage As Double

        If RightDistance <> LeftDistance Then
            InterpolatedPercentage = (unitsX - LeftDistance) / (RightDistance - LeftDistance)
        Else
            InterpolatedPercentage = 0
        End If

        Dim lp As bbPoint = StaffPathPoints(LeftPoint)
        Dim rp As bbPoint = StaffPathPoints(RightPoint)

        'The second thing to do is to interpolate the X and Y coordinates

        Dim lx As Double = lp.X
        Dim ly As Double = lp.Y
        Dim rx As Double = rp.X
        Dim ry As Double = rp.Y

        InterpolatedPoint.X = lx + (rx - lx) * InterpolatedPercentage
        InterpolatedPoint.Y = ly + (ry - ly) * InterpolatedPercentage
        unitsSpaceHeight = (LeftStaffHeight + (RightStaffHeight - LeftStaffHeight) * InterpolatedPercentage)

        Dim spacesYPosition As Double = 0.5 * spacesY * unitsSpaceHeight
        Dim staffLinePosition As bbPoint = FindDisplacement(InterpolatedPoint, rp, spacesYPosition)
        Dim staffLineOrientation As New bbRay(InterpolatedPoint, rp)

        'Make the angle perpendicular in the direction of the stem
        unitsOrientation = New bbRay(staffLinePosition, staffLinePosition)
        unitsOrientation.angle = staffLineOrientation.angle

        FindPlottingPointTime.StopClock()
    End Sub

    Public Sub FindPlottingPoint(ByVal unitsX As Double, ByVal spacesY As Double, ByRef unitsOrientation As bbRay, ByRef unitsSpaceHeight As Double)
        FindPlottingPoint(unitsX, spacesY, unitsOrientation, unitsSpaceHeight, 0)
    End Sub

    Public Sub FindPlottingPointEx(ByVal unitsX As Double, ByVal spacesY As Double, ByRef unitsOrientation As bbRay, ByRef unitsSpaceHeight As Double, ByRef LeftStaffPoint As Integer, ByVal KernScale As Double)
        FindPlottingPointTime.StartClock()
        'The first thing to do is find between which two points the horizontal component exists
        Dim info As bbStaffPointInfo = FindPoint(unitsX)

        Dim LeftPoint As Integer = info.LeftPoint
        Dim RightPoint As Integer = info.RightPoint
        Dim LeftDistance As Double = info.LeftDistance
        Dim RightDistance As Double = info.RightDistance
        Dim LeftStaffHeight As Double = info.LeftStaffHeight.Value
        Dim RightStaffHeight As Double = info.RightStaffHeight.Value
        LeftStaffPoint = LeftPoint
        Dim InterpolatedPoint As New bbPoint
        Dim InterpolatedPercentage As Double

        If RightDistance <> LeftDistance Then
            InterpolatedPercentage = (unitsX - LeftDistance) / (RightDistance - LeftDistance)
        Else
            InterpolatedPercentage = 0
        End If

        InterpolatedPercentage = InterpolatedPercentage - 1 / 2
        InterpolatedPercentage *= KernScale
        InterpolatedPercentage = InterpolatedPercentage + 1 / 2 * KernScale


        Dim lp As bbPoint = StaffPathPoints(LeftPoint)
        Dim rp As bbPoint = StaffPathPoints(RightPoint)
        Dim ip As New bbPoint

        'The second thing to do is to interpolate the X and Y coordinates

        Dim lx As Double = lp.X
        Dim ly As Double = lp.Y
        Dim rx As Double = rp.X
        Dim ry As Double = rp.Y



        InterpolatedPoint.X = lx + (rx - lx) * InterpolatedPercentage
        InterpolatedPoint.Y = ly + (ry - ly) * InterpolatedPercentage
        ip.X = lx + (rx - lx) * InterpolatedPercentage * 2
        ip.Y = ly + (ry - ly) * InterpolatedPercentage * 2

        unitsSpaceHeight = (LeftStaffHeight + (RightStaffHeight - LeftStaffHeight) * InterpolatedPercentage)

        Dim spacesYPosition As Double = 0.5 * spacesY * unitsSpaceHeight

        Dim staffLinePosition As New bbPoint

        Dim staffLineOrientation As bbRay
        If InterpolatedPercentage >= 1 Then
            staffLinePosition = FindDisplacement(InterpolatedPoint, ip, spacesYPosition)
            staffLineOrientation = New bbRay(InterpolatedPoint, ip)
        Else
            staffLinePosition = FindDisplacement(InterpolatedPoint, rp, spacesYPosition)
            staffLineOrientation = New bbRay(InterpolatedPoint, rp)
        End If



        'Make the angle perpendicular in the direction of the stem
        unitsOrientation = New bbRay(staffLinePosition, staffLinePosition)
        unitsOrientation.angle = staffLineOrientation.angle

        FindPlottingPointTime.StopClock()
    End Sub

    Public Property StaffProperty(ByVal PathPointIndex As Integer) As bbStaffProperties
        Get
            'Force data at Index to exist in case Get is called to obtain a reference
            'to internal members. This is a public method intended to help
            'outside commands get access to select information quickly...
            Dim x As bbStaffProperties = StaffProperties(PathPointIndex)
            If IsNothing(x) Then
                'We have to return something, and we have to make 
                'sure it gets added to the collection
                x = New bbStaffProperties
                StaffProperties(PathPointIndex) = x
            End If
            Return x
        End Get
        Set(ByVal value As bbStaffProperties)
            StaffProperties(PathPointIndex) = value
        End Set
    End Property

    Public Function AbsoluteHighestStaffLine() As Integer
        If Not HighestLinePrecomputed Then
            Dim Maximum As Integer = Integer.MinValue
            For i As Integer = StaffPathPoints.LBound() To StaffPathPoints.UBound()
                Dim x As bbStaffProperties = StaffProperties(i)
                If Not IsNothing(x) Then
                    Maximum = Math.Max(x.LinesUBound(), Maximum)
                End If
            Next
            HighestLine = Maximum
            HighestLinePrecomputed = True
        End If
        Return HighestLine
    End Function

    Public Function AbsoluteLowestStaffLine() As Integer
        If Not LowestLinePrecomputed Then
            Dim Minimum As Integer = Integer.MaxValue
            For i As Integer = StaffPathPoints.LBound() To StaffPathPoints.UBound()
                Dim x As bbStaffProperties = StaffProperties(i)
                If Not IsNothing(x) Then
                    Minimum = Math.Min(x.LinesLBound(), Minimum)
                End If
            Next
            LowestLine = Minimum
            LowestLinePrecomputed = True
        End If
        Return LowestLine
    End Function

    Public Sub DrawStaff()
        'Make sure there are enough points to trace lines!
        If StaffPathPoints.Count() < 2 Then Exit Sub

        Parent.AddLineToStream(SaveCTM())

        'Set default staff properties
        Parent.AddLineToStream(PathJoin(True))

        Dim DefaultStaffProperties As New bbStaffProperties

        With DefaultStaffProperties
            .SpaceHeight.Value = 0.0 'We should not set a default for this property

            For eachStaffLine As Integer = AbsoluteLowestStaffLine() To AbsoluteHighestStaffLine()
                With .StaffLine(eachStaffLine)
                    .Draw.Value = True
                    .LineThickness.Value = 0.01
                    .DrawDashedLine.Value = False
                    .DashOnLength.Value = 0.1
                    .DashOffLength.Value = 0.1
                    .DashOffset.Value = 0
                    .DrawFlatCap.Value = False
                End With
            Next eachStaffLine
        End With

        For eachStaffLine As Integer = AbsoluteLowestStaffLine() To AbsoluteHighestStaffLine()
            'Loop through each possible staff line....
            'Save the current state, so the other lines can be drawn independently
            Parent.AddLineToStream(SaveCTM())

            'Set the current staff properties anew (MAKE SURE TO UPDATE THIS)

            Dim CurrentStaffProperties As New bbStaffProperties
            Dim PreviousStaffProperties As New bbStaffProperties

            CurrentStaffProperties.SpaceHeight.Value = DefaultStaffProperties.SpaceHeight.Value
            CurrentStaffProperties.StaffLine(eachStaffLine) += DefaultStaffProperties.StaffLine(eachStaffLine)

            Dim WasDrawingPath As Boolean = False 'Indicates whether the first point has been plotted
            Dim WillNotBeDrawingPath As Boolean = False 'Indicates that the last point is being plotted
            Dim MustRestartPath As Boolean = False 'Indicates when properties change and cause the path to be restarted

            For eachPathPoint As Integer = StaffPathPoints.LBound() To StaffPathPoints.UBound()
                If eachPathPoint = StaffPathPoints.UBound() Then
                    'Last point has been reached
                    WillNotBeDrawingPath = True
                End If
                'Update properties
                CurrentStaffProperties.SpaceHeight += StaffProperty(eachPathPoint).SpaceHeight
                CurrentStaffProperties.StaffLine(eachStaffLine) += StaffProperty(eachPathPoint).StaffLine(eachStaffLine)
                'PreviousStaffProperties.SpaceHeight += StaffProperty(eachPathPoint).SpaceHeight
                'PreviousStaffProperties.StaffLine(eachStaffLine) += StaffProperty(eachPathPoint).StaffLine(eachStaffLine)

                'See if the path needs to be restarted due to property changes
                If CurrentStaffProperties.StaffLine(eachStaffLine).Changed() Then
                    If WasDrawingPath Then
                        MustRestartPath = True
                    End If
                End If

                'Determine where the point is to be drawn
                Dim p As bbPoint
                
                If eachPathPoint = StaffPathPoints.LBound() Then
                    Dim p1 = Point(0)
                    Dim p2 = Point(1)
                    Dim d1 = CurrentStaffProperties.SpaceHeight.Value * eachStaffLine
                    'Left point (use first two points to calculate right angle)
                    p = FindDisplacement(p1, p2, d1)
                ElseIf eachPathPoint = StaffPathPoints.UBound() Then
                    Dim p1 = Point(StaffPathPoints.UBound())
                    Dim p2 = Point(StaffPathPoints.UBound() - 1)
                    Dim d1 = -CurrentStaffProperties.SpaceHeight.Value * eachStaffLine
                    'Right point (use last two points to calculate right angle)
                    p = FindDisplacement(p1, p2, d1)
                Else
                    Dim p1 = Point(eachPathPoint - 1)
                    Dim p2 = Point(eachPathPoint)
                    Dim p3 = Point(eachPathPoint + 1)
                    Dim d1 = CurrentStaffProperties.SpaceHeight.Value * eachStaffLine
                    Dim d2 = CurrentStaffProperties.SpaceHeight.Value * eachStaffLine
                    'All other middle points (use three points to calculate bisecting angle)
                    p = FindIntersectionAfterDisplacement(p1, p2, p3, d1, d2)
                End If

                If Not WasDrawingPath Then
                    'Set the path drawing properties
                    Parent.AddLineToStream(CurrentStaffProperties.StaffLine(eachStaffLine).SetProperties())

                    'Start drawing the path anew
                    Parent.AddLineToStream(PathStart(p.X, p.Y))

                    WasDrawingPath = True
                    MustRestartPath = False
                Else
                    If WillNotBeDrawingPath Then
                        'Add the last point to the current path line
                        Parent.AddLineToStream(PathContinue(p.X, p.Y))

                        'End the old path, but check to make sure it's not a "no-op" no-draw
                        If CurrentStaffProperties.StaffLine(eachStaffLine).Draw.Value = True Then
                            Parent.AddLineToStream(DrawPath())
                        Else
                            Parent.AddLineToStream(DoNotDrawPath())
                        End If
                    ElseIf MustRestartPath = False Then
                        'Simply continue drawing the path
                        Parent.AddLineToStream(PathContinue(p.X, p.Y))
                    Else
                        'End the old path, and begin a new path on the same point

                        'Add the last point to the current path line
                        Parent.AddLineToStream(PathContinue(p.X, p.Y))

                        'End the old path, but check to make sure it's not a "no-op" no-draw
                        If PreviousStaffProperties.StaffLine(eachStaffLine).Draw.Value = True Then
                            Parent.AddLineToStream(DrawPath())
                        Else
                            Parent.AddLineToStream(DoNotDrawPath())
                        End If

                        'Set the properties
                        Parent.AddLineToStream(CurrentStaffProperties.StaffLine(eachStaffLine).SetProperties())

                        'Begin the new path
                        Parent.AddLineToStream(PathStart(p.X, p.Y))
                        MustRestartPath = False
                    End If
                End If

                'Update the "previous" properties to reflect the current properties
                PreviousStaffProperties.SpaceHeight.Value = CurrentStaffProperties.SpaceHeight.Value
                PreviousStaffProperties.StaffLine(eachStaffLine) += CurrentStaffProperties.StaffLine(eachStaffLine)

            Next
            Parent.AddLineToStream(RevertCTM())
        Next

        'Undo the CTM changes we made, so that other
        'staves can be processed independently
        Parent.AddLineToStream(RevertCTM())
    End Sub

    Private Sub Z_Test_MeasureLines()
        'Try to draw some test measure lines
        'For i As Integer = 10 To Points() - 10 Step 40
        'Dim p1 As bbPoint = FindIntersectionAfterDisplacement(Point(i - 1), Point(i), Point(i + 1), -LineSpace * LinesBelow, -LineSpace * LinesBelow)
        'Dim p2 As bbPoint = FindIntersectionAfterDisplacement(Point(i - 1), Point(i), Point(i + 1), LineSpace * LinesAbove, LineSpace * LinesAbove)
        'Parent.AddLineToStream(PathStart(p1.X, p1.Y))
        'Parent.AddLineToStream(PathContinue(p2.X, p2.Y))
        'Parent.AddLineToStream(DrawPath())
        'Next
    End Sub
End Class
