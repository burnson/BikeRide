Public Module bbNote
    Public Sub DrawBasicEighthNotes(ByRef Staff As bbStaff, ByRef NoteGroup As Collection, ByRef AccidentalFont As bbFont)
        Dim l As bbNoteProperty = NoteGroup.Item(1)
        Dim r As bbNoteProperty = NoteGroup.Item(NoteGroup.Count)
        For i As Integer = 1 To NoteGroup.Count()
            Dim n As bbNoteProperty = NoteGroup.Item(i)
            DrawNoteWithStem(Staff, n.unitsX, n.NoteName, True, n.StemUp, n.StemWidth, n.StemHeight, n.Color, n.Accidental, AccidentalFont)
        Next
        Dim BeamThickness As Double = Globals.ratioBeamThicknessToSpaceHeight
        If l.StemUp Then
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight, l.StemWidth, r.unitsX, r.NoteName + r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        Else
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight, l.StemWidth, r.unitsX, r.NoteName - r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        End If
    End Sub

    Public Sub DrawRhythm2(ByRef Staff As bbStaff, ByRef NoteGroup As Collection, ByRef AccidentalFont As bbFont)
        Dim l As bbNoteProperty = NoteGroup.Item(1)
        Dim l2 As bbNoteProperty = NoteGroup.Item(2)
        Dim r As bbNoteProperty = NoteGroup.Item(NoteGroup.Count)
        For i As Integer = 1 To NoteGroup.Count()
            Dim n As bbNoteProperty = NoteGroup.Item(i)
            DrawNoteWithStem(Staff, n.unitsX, n.NoteName, True, n.StemUp, n.StemWidth, n.StemHeight, n.Color, n.Accidental, AccidentalFont)
        Next
        Dim BeamThickness As Double = Globals.ratioBeamThicknessToSpaceHeight
        If l.StemUp Then
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight, l.StemWidth, r.unitsX, r.NoteName + r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight - Globals.spacesBeamSeparation, l.StemWidth, l2.unitsX, l2.NoteName + l2.StemHeight - Globals.spacesBeamSeparation, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        Else
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight, l.StemWidth, r.unitsX, r.NoteName - r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight + Globals.spacesBeamSeparation, l.StemWidth, l2.unitsX, l2.NoteName - l2.StemHeight + Globals.spacesBeamSeparation, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        End If
    End Sub

    Public Sub DrawRhythm3(ByRef Staff As bbStaff, ByRef NoteGroup As Collection, ByRef AccidentalFont As bbFont)
        Dim l As bbNoteProperty = NoteGroup.Item(1)
        Dim l2 As bbNoteProperty = NoteGroup.Item(NoteGroup.Count)
        Dim r As bbNoteProperty = NoteGroup.Item(NoteGroup.Count)
        For i As Integer = 1 To NoteGroup.Count()
            Dim n As bbNoteProperty = NoteGroup.Item(i)
            DrawNoteWithStem(Staff, n.unitsX, n.NoteName, True, n.StemUp, n.StemWidth, n.StemHeight, n.Color, n.Accidental, AccidentalFont)
        Next
        Dim BeamThickness As Double = Globals.ratioBeamThicknessToSpaceHeight
        If l.StemUp Then
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight, l.StemWidth, r.unitsX, r.NoteName + r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight - Globals.spacesBeamSeparation, l.StemWidth, l2.unitsX, l2.NoteName + l2.StemHeight - Globals.spacesBeamSeparation, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        Else
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight, l.StemWidth, r.unitsX, r.NoteName - r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight + Globals.spacesBeamSeparation, l.StemWidth, l2.unitsX, l2.NoteName - l2.StemHeight + Globals.spacesBeamSeparation, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        End If
    End Sub

    Public Sub DrawRhythm4(ByRef Staff As bbStaff, ByRef NoteGroup As Collection, ByRef AccidentalFont As bbFont)
        Dim l As bbNoteProperty = NoteGroup.Item(1)
        Dim l2 As bbNoteProperty = NoteGroup.Item(NoteGroup.Count)
        Dim r As bbNoteProperty = NoteGroup.Item(NoteGroup.Count)
        For i As Integer = 1 To NoteGroup.Count()
            Dim n As bbNoteProperty = NoteGroup.Item(i)
            DrawNoteWithStem(Staff, n.unitsX, n.NoteName, True, n.StemUp, n.StemWidth, n.StemHeight, n.Color, n.Accidental, AccidentalFont)
        Next
        Dim BeamThickness As Double = Globals.ratioBeamThicknessToSpaceHeight
        If l.StemUp Then
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight, l.StemWidth, r.unitsX, r.NoteName + r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight - Globals.spacesBeamSeparation, l.StemWidth, l2.unitsX, l2.NoteName + l2.StemHeight - Globals.spacesBeamSeparation, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName + l.StemHeight - Globals.spacesBeamSeparation * 2.0, l.StemWidth, l2.unitsX, l2.NoteName + l2.StemHeight - Globals.spacesBeamSeparation * 2.0, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        Else
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight, l.StemWidth, r.unitsX, r.NoteName - r.StemHeight, r.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight + Globals.spacesBeamSeparation, l.StemWidth, l2.unitsX, l2.NoteName - l2.StemHeight + Globals.spacesBeamSeparation, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
            DrawBeam(Staff, l.unitsX, l.NoteName - l.StemHeight + Globals.spacesBeamSeparation * 2.0, l.StemWidth, l2.unitsX, l2.NoteName - l2.StemHeight + Globals.spacesBeamSeparation * 2.0, l2.StemWidth, BeamThickness, BeamThickness, Globals.numberBeamInterpolations, l.Color)
        End If
    End Sub
    Public Sub SetBasicNoteGroupStemHeights(ByRef NoteGroup As Collection)
        SetBasicNoteGroupStemHeights(NoteGroup, Globals.spacesDefaultStemHeight, Globals.spacesDefaultStemHeight)
    End Sub

    Public Sub SetBasicNoteGroupStemHeights(ByRef NoteGroup As Collection, ByVal spacesLeftHeight As Double, ByVal spacesRightHeight As Double)
        Dim LeftNote As bbNoteProperty = NoteGroup.Item(1)
        Dim RightNote As bbNoteProperty = NoteGroup.Item(NoteGroup.Count())

        Dim LeftY As Double
        Dim RightY As Double

        If LeftNote.StemUp Then
            LeftY = LeftNote.NoteName + spacesLeftHeight
            RightY = RightNote.NoteName + spacesRightHeight
        Else
            LeftY = LeftNote.NoteName - spacesLeftHeight
            RightY = RightNote.NoteName - spacesRightHeight
        End If

        For i As Integer = 1 To NoteGroup.Count
            Dim p As Double = 0
            If i > 1 Then
                p = (i - 1) / (NoteGroup.Count - 1)
            End If
            Dim InterpolatedY As Double = (RightY - LeftY) * p + LeftY
            Dim InterpolatedNote As bbNoteProperty = NoteGroup.Item(i)
            Dim InterpolatedSpaces As Double = Math.Abs(InterpolatedY - InterpolatedNote.NoteName)
            InterpolatedNote.StemHeight = InterpolatedSpaces
        Next
    End Sub

    Public Sub SetBasicNoteGroupStemDirections(ByRef NoteGroup As Collection)
        Dim Sum As Integer = 0
        For i As Integer = 1 To NoteGroup.Count
            Dim CurrentNote As bbNoteProperty = NoteGroup.Item(i)
            Sum += CurrentNote.NoteName
        Next
        Dim DirectionUp As Boolean
        If Sum >= 0 Then
            DirectionUp = False
        Else
            DirectionUp = True
        End If
        For i As Integer = 1 To NoteGroup.Count
            Dim CurrentNote As bbNoteProperty = NoteGroup.Item(i)
            CurrentNote.StemUp = DirectionUp
        Next
    End Sub

    Public Sub DrawNoteWithStemAndFlag(ByRef Staff As bbStaff, ByVal unitsX As Double, _
                                ByVal spacesY As Double, ByVal NoteheadFilled As Boolean, _
                                ByVal StemUp As Boolean, ByVal stemWidth As Double, _
                                ByVal spacesStemHeight As Double, ByVal Color As bbColor, _
                                ByVal Accidental As bbAccidentalType, ByRef AccidentalFont As bbFont)

        DrawNoteWithStemTime.StartClock()

        Dim Page As bbPDFPage = Nothing
        Staff.GetParent(Page)

        'Draw some ledger lines -- time is about 4% of the 10%
        Dim HorizontalOrientation As bbRay = Nothing
        Dim LocalSpaceHeight As Double = 0.0
        Dim StaffPoint As Integer = 0
        Dim LedgerRangeLeft As Integer = 0
        Dim LedgerRangeRight As Integer = 0
        Dim LedgerStep As Integer = 0
        Dim LedgerLeft As bbRay = Nothing
        Dim LedgerRight As bbRay = Nothing
        Dim HalfLedgerWidth As Double = 0.0
        Dim HalfNoteWidth As Double = 0.0
        Dim NoteCenter As Double = 0.0
        Dim p1 As bbPoint = Nothing
        Dim p2 As bbPoint = Nothing
        Dim i As Integer = 0

        Staff.FindPlottingPoint(unitsX, spacesY, HorizontalOrientation, LocalSpaceHeight, StaffPoint)
        HalfLedgerWidth = LocalSpaceHeight * Globals.ratioLedgerLineToSpaceHeight / 2.0
        HalfNoteWidth = GetNoteWidth(LocalSpaceHeight) / 2.0

        LedgerRangeRight = CLng(Fix(spacesY))

        If spacesY > 0 Then
            LedgerRangeLeft = 2 * (Staff.AbsoluteHighestStaffLine() + 1)
            LedgerStep = 2
        Else
            LedgerRangeLeft = 2 * (Staff.AbsoluteLowestStaffLine() - 1)
            LedgerStep = -2
        End If

        Page.AddLineToStream(SaveCTM())
        Page.AddLineToStream(PathCap(True))
        For i = LedgerRangeLeft To LedgerRangeRight Step LedgerStep
            If StemUp Then
                NoteCenter = unitsX - HalfNoteWidth
            Else
                NoteCenter = unitsX + HalfNoteWidth
            End If

            Staff.FindPlottingPoint(NoteCenter - HalfLedgerWidth, i, LedgerLeft, 0)
            Staff.FindPlottingPoint(NoteCenter + HalfLedgerWidth, i, LedgerRight, 0)
            p1 = New bbPoint(LedgerLeft.x, LedgerLeft.y)
            p2 = New bbPoint(LedgerRight.x, LedgerRight.y)

            With Staff.StaffProperty(0).StaffLine(Staff.AbsoluteHighestStaffLine())
                Page.AddLineToStream(DrawLine(p1, p2, .LineThickness.Value, .LineStrokeColor.Value))
            End With
        Next
        Page.AddLineToStream(RevertCTM())

        'Draw the note 2.34% out of 10% of the total time

        Dim StemTrim As Double = 0.0

        Staff.FindPlottingPoint(unitsX, spacesY, HorizontalOrientation, LocalSpaceHeight)

        If NoteheadFilled Then
            Page.AddLineToStream(QuarterNote(New bbPoint(HorizontalOrientation.x, _
            HorizontalOrientation.y), stemWidth, StemUp, HorizontalOrientation.angle, _
            LocalSpaceHeight, 1.0, Color, StemTrim))
        Else
            Page.AddLineToStream(HalfNote(New bbPoint(HorizontalOrientation.x, _
            HorizontalOrientation.y), stemWidth, StemUp, HorizontalOrientation.angle, _
            LocalSpaceHeight, 1.0, Color, StemTrim))
        End If


        'The following took about 1.6% out of 10% of the total time *****************
        'Do accidental
        Dim AccidentalText As String = ""
        Select Case Accidental
            Case bbAccidentalType.None
                AccidentalText = ""
            Case bbAccidentalType.Natural
                AccidentalText = "a"
            Case bbAccidentalType.Sharp
                AccidentalText = "b"
            Case bbAccidentalType.Flat
                AccidentalText = "c"
            Case bbAccidentalType.NaturalPairUp
                AccidentalText = "A"
            Case bbAccidentalType.SharpPairUp
                AccidentalText = "B"
            Case bbAccidentalType.FlatPairUp
                AccidentalText = "C"
        End Select
        Dim AccidentalPosition As bbRay = Nothing
        Dim AccidentalNoteWidth As Double
        If StemUp Then
            AccidentalNoteWidth = GetNoteWidth(LocalSpaceHeight) * 2 - stemWidth
        Else
            AccidentalNoteWidth = 0
        End If
        Staff.FindPlottingPoint(unitsX - AccidentalNoteWidth - LocalSpaceHeight * Globals.ratioAccidentalDisplacementToSpaceHeight, spacesY, AccidentalPosition, 0)
        Page.AddLineToStream(DrawString(AccidentalFont, LocalSpaceHeight, AccidentalPosition, AccidentalText))

        'The following is PRETTY FAST*************

        'Do stem with flag
        Dim HorizontalTestPoint1 As New bbPoint(HorizontalOrientation.x, HorizontalOrientation.y)
        Dim HorizontalTestPoint2 As New bbPoint(HorizontalOrientation.x + Math.Cos(HorizontalOrientation.angle), HorizontalOrientation.y + Math.Sin(HorizontalOrientation.angle))
        Dim StemBottomPoint As bbPoint, StemTopPoint As bbPoint

        StemBottomPoint = FindDisplacement(HorizontalTestPoint1, HorizontalTestPoint2, StemTrim)
        Dim Direction As Double
        If StemUp Then
            'Divide by two because spaces refer to the vertical height sequence:
            'space...line...space..line and not space...space...space
            Direction = 1.0 / 2.0
        Else
            Direction = -1.0 / 2.0
        End If
        StemTopPoint = FindDisplacement(HorizontalTestPoint1, HorizontalTestPoint2, Direction * spacesStemHeight * LocalSpaceHeight)

        Dim fudge_value As Double = 0.98 'prevents a very thin line from not getting drawn in between the stem and the flag
        Page.AddLineToStream(SaveCTM())
        Page.AddLineToStream(PathCap(True))
        Page.AddLineToStream(DrawLine(StemBottomPoint, StemTopPoint, stemWidth, Color))
        Dim FlagOrientation As New bbRay(StemTopPoint.X + Math.Cos(HorizontalOrientation.angle) * stemWidth / 2 * fudge_value, StemTopPoint.Y + Math.Sin(HorizontalOrientation.angle) * stemWidth / 2 * fudge_value, HorizontalOrientation.angle)
        Page.AddLineToStream(DrawString(NotationFont, LocalSpaceHeight, FlagOrientation, "J"))
        Page.AddLineToStream(RevertCTM())
        DrawNoteWithStemTime.StopClock()
    End Sub

    Public Sub DrawNoteWithStem(ByRef Staff As bbStaff, ByVal unitsX As Double, _
                                ByVal spacesY As Double, ByVal NoteheadFilled As Boolean, _
                                ByVal StemUp As Boolean, ByVal stemWidth As Double, _
                                ByVal spacesStemHeight As Double, ByVal Color As bbColor, _
                                ByVal Accidental As bbAccidentalType, ByRef AccidentalFont As bbFont)

        DrawNoteWithStemTime.StartClock()

        Dim Page As bbPDFPage = Nothing
        Staff.GetParent(Page)

        'Draw some ledger lines -- time is about 4% of the 10%
        Dim HorizontalOrientation As bbRay = Nothing
        Dim LocalSpaceHeight As Double = 0.0
        Dim StaffPoint As Integer = 0
        Dim LedgerRangeLeft As Integer = 0
        Dim LedgerRangeRight As Integer = 0
        Dim LedgerStep As Integer = 0
        Dim LedgerLeft As bbRay = Nothing
        Dim LedgerRight As bbRay = Nothing
        Dim HalfLedgerWidth As Double = 0.0
        Dim HalfNoteWidth As Double = 0.0
        Dim NoteCenter As Double = 0.0
        Dim p1 As bbPoint = Nothing
        Dim p2 As bbPoint = Nothing
        Dim i As Integer = 0

        Staff.FindPlottingPoint(unitsX, spacesY, HorizontalOrientation, LocalSpaceHeight, StaffPoint)
        HalfLedgerWidth = LocalSpaceHeight * Globals.ratioLedgerLineToSpaceHeight / 2.0
        HalfNoteWidth = GetNoteWidth(LocalSpaceHeight) / 2.0

        LedgerRangeRight = CLng(Fix(spacesY))

        If spacesY > 0 Then
            LedgerRangeLeft = 2 * (Staff.AbsoluteHighestStaffLine() + 1)
            LedgerStep = 2
        Else
            LedgerRangeLeft = 2 * (Staff.AbsoluteLowestStaffLine() - 1)
            LedgerStep = -2
        End If

        Page.AddLineToStream(SaveCTM())
        Page.AddLineToStream(PathCap(True))
        For i = LedgerRangeLeft To LedgerRangeRight Step LedgerStep
            If StemUp Then
                NoteCenter = unitsX - HalfNoteWidth
            Else
                NoteCenter = unitsX + HalfNoteWidth
            End If

            Staff.FindPlottingPoint(NoteCenter - HalfLedgerWidth, i, LedgerLeft, 0)
            Staff.FindPlottingPoint(NoteCenter + HalfLedgerWidth, i, LedgerRight, 0)
            p1 = New bbPoint(LedgerLeft.x, LedgerLeft.y)
            p2 = New bbPoint(LedgerRight.x, LedgerRight.y)

            With Staff.StaffProperty(0).StaffLine(Staff.AbsoluteHighestStaffLine())
                Page.AddLineToStream(DrawLine(p1, p2, .LineThickness.Value, .LineStrokeColor.Value))
            End With
        Next
        Page.AddLineToStream(RevertCTM())

        'Draw the note 2.34% out of 10% of the total time

        Dim StemTrim As Double = 0.0

        Staff.FindPlottingPoint(unitsX, spacesY, HorizontalOrientation, LocalSpaceHeight)

        If NoteheadFilled Then
            Page.AddLineToStream(QuarterNote(New bbPoint(HorizontalOrientation.x, _
            HorizontalOrientation.y), stemWidth, StemUp, HorizontalOrientation.angle, _
            LocalSpaceHeight, 1.0, Color, StemTrim))
        Else
            Page.AddLineToStream(HalfNote(New bbPoint(HorizontalOrientation.x, _
            HorizontalOrientation.y), stemWidth, StemUp, HorizontalOrientation.angle, _
            LocalSpaceHeight, 1.0, Color, StemTrim))
        End If


        'The following took about 1.6% out of 10% of the total time *****************
        'Do accidental
        Dim AccidentalText As String = ""
        Select Case Accidental
            Case bbAccidentalType.None
                AccidentalText = ""
            Case bbAccidentalType.Natural
                AccidentalText = "a"
            Case bbAccidentalType.Sharp
                AccidentalText = "b"
            Case bbAccidentalType.Flat
                AccidentalText = "c"
            Case bbAccidentalType.NaturalPairUp
                AccidentalText = "A"
            Case bbAccidentalType.SharpPairUp
                AccidentalText = "B"
            Case bbAccidentalType.FlatPairUp
                AccidentalText = "C"
        End Select
        Dim AccidentalPosition As bbRay = Nothing
        Dim AccidentalNoteWidth As Double
        If StemUp Then
            AccidentalNoteWidth = GetNoteWidth(LocalSpaceHeight) * 2 - stemWidth
        Else
            AccidentalNoteWidth = 0
        End If
        Staff.FindPlottingPoint(unitsX - AccidentalNoteWidth - LocalSpaceHeight * Globals.ratioAccidentalDisplacementToSpaceHeight, spacesY, AccidentalPosition, 0)
        Page.AddLineToStream(DrawString(AccidentalFont, LocalSpaceHeight, AccidentalPosition, AccidentalText))

        'The following is PRETTY FAST*************
        'Do stem

        Dim HorizontalTestPoint1 As New bbPoint(HorizontalOrientation.x, HorizontalOrientation.y)
        Dim HorizontalTestPoint2 As New bbPoint(HorizontalOrientation.x + Math.Cos(HorizontalOrientation.angle), HorizontalOrientation.y + Math.Sin(HorizontalOrientation.angle))
        Dim StemBottomPoint As bbPoint, StemTopPoint As bbPoint

        StemBottomPoint = FindDisplacement(HorizontalTestPoint1, HorizontalTestPoint2, StemTrim)
        Dim Direction As Double
        If StemUp Then
            'Divide by two because spaces refer to the vertical height sequence:
            'space...line...space..line and not space...space...space
            Direction = 1.0 / 2.0
        Else
            Direction = -1.0 / 2.0
        End If
        StemTopPoint = FindDisplacement(HorizontalTestPoint1, HorizontalTestPoint2, Direction * spacesStemHeight * LocalSpaceHeight)

        Page.AddLineToStream(SaveCTM())
        Page.AddLineToStream(PathCap(True))
        Page.AddLineToStream(DrawLine(StemBottomPoint, StemTopPoint, stemWidth, Color))
        Page.AddLineToStream(RevertCTM())


        DrawNoteWithStemTime.StopClock()
    End Sub

End Module
