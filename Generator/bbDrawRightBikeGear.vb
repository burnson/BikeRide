Module bbDrawRightBikeGear
    Dim Page As bbPDFPage = Nothing
    Dim Staff As bbStaff
    Dim l As Double
    Dim w As Double
    Dim c As New bbColor(0, 0, 0)
    Dim o As Double
    Dim t As Double = StemThickness
    Dim n As Double
    Dim m As Collection

    Const nat As Integer = bbAccidentalType.Natural
    Const shp As Integer = bbAccidentalType.Sharp
    Const flt As Integer = bbAccidentalType.Flat
    Const nne As Integer = bbAccidentalType.None
    Const pnat As Integer = bbAccidentalType.NaturalPairUp
    Const pshp As Integer = bbAccidentalType.SharpPairUp
    Const pflt As Integer = bbAccidentalType.FlatPairUp

    Private Sub DrawCresc(ByVal IsCresc As Boolean, ByVal LeftMeasure As Double, ByVal relativeXLeft As Double, ByVal relativeXLeftWidth As Double, ByVal RightMeasure As Double, ByVal relativeXRight As Double, ByVal relativeXRightWidth As Double, ByVal spacesY As Double, ByVal apex As Double)
        Dim x1 As Double = GetMeasuredX(LeftMeasure, relativeXLeft, relativeXLeftWidth)
        Dim x2 As Double = GetMeasuredX(RightMeasure, relativeXRight, relativeXRightWidth)
        DrawHairpin(Staff, IsCresc, x1, spacesY, x2, spacesY, apex, Globals.ratioStaffLineToSpaceHeight, Globals.numberBeamInterpolations, New bbColor(0, 0, 0))
    End Sub

    Private Function GetMeasuredX(ByVal Measure As Double, ByVal relativeX As Double, ByVal relativeWidth As Double) As Double
        PrepareMeasure(Measure)
        Return (relativeX + 0.5) / relativeWidth * w + o
    End Function

    Private Sub PrepareMeasure(ByVal Number As Double)
        n = Number * 2 - 1
        m = New Collection
        o = l * (n - 1) + l * Globals.percentageMeasureLeftMargin
    End Sub

    Private Sub WriteNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal noteName As bbNoteName, ByVal noteAccidental As bbAccidentalType)
        m.Add(New bbNoteProperty((relativeX + 0.5) / (relativeWidth), noteName, noteAccidental, w, o, c, t))
    End Sub

    Private Sub WriteSymbol(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal spacesY As Double, ByVal symbol As String)
        Dim LocalSpaceHeight As Double
        Dim Orientation As bbRay = Nothing
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        Staff.FindPlottingPoint(x, spacesY, Orientation, LocalSpaceHeight, 0)
        Page.AddLineToStream(DrawString(NotationFont, LocalSpaceHeight, Orientation, symbol))
    End Sub

    Private Sub WriteQuarterNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal NoteName As bbNoteName, ByVal NoteAccidental As bbAccidentalType)
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        DrawNoteWithStem(Staff, x, NoteName, True, True, StemThickness, Globals.spacesDefaultStemHeight, c, NoteAccidental, NotationFont)
    End Sub

    Private Sub WriteHalfNoteUp(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal NoteName As bbNoteName, ByVal NoteAccidental As bbAccidentalType)
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        DrawNoteWithStem(Staff, x, NoteName, False, True, StemThickness, Globals.spacesDefaultStemHeight, c, NoteAccidental, NotationFont)
    End Sub


    Private Sub WriteHalfNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal NoteName As bbNoteName, ByVal NoteAccidental As bbAccidentalType)
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        Dim LocalSpaceHeight As Double
        Dim HorizontalOrientation As bbRay = Nothing
        Staff.FindPlottingPoint(x, 0, HorizontalOrientation, LocalSpaceHeight, 0)
        x -= (GetNoteWidth(LocalSpaceHeight) + StemThickness / 2)
        DrawNoteWithStem(Staff, x, NoteName, False, False, StemThickness, Globals.spacesDefaultStemHeight, c, NoteAccidental, NotationFont)
    End Sub

    Private Sub WriteEighthRest(ByVal relativeX As Double, ByVal relativeWidth As Double)
        WriteSymbol(relativeX, relativeWidth, 0, "h")
    End Sub

    Private Sub DrawNotesUp()
        For i As Integer = 1 To m.Count
            Dim x As bbNoteProperty = m.Item(i)
            x.StemUp = True
        Next
        SetBasicNoteGroupStemHeights(m)
        DrawBasicEighthNotes(Staff, m, NotationFont)
    End Sub

    Private Sub DrawNotes(ByVal LeftHeight As Double, ByVal RightHeight As Double)
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m, LeftHeight, RightHeight)
        DrawBasicEighthNotes(Staff, m, NotationFont)
    End Sub

    Private Sub DrawNotes()
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m)
        DrawBasicEighthNotes(Staff, m, NotationFont)
    End Sub

    Public Sub DrawRightBikeGear(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 12
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'Ostinato
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0, 4.8, bbNoteName.TrebleG, "d")
        WriteNote(0.8, 4.8, bbNoteName.TrebleBUp, flt)
        WriteNote(1.8, 4.8, bbNoteName.TrebleAUp, flt)
        WriteNote(2.8, 4.8, bbNoteName.TrebleBUp, nne)
        WriteNote(3.8, 4.8, bbNoteName.TrebleCUp, nne)
        DrawNotesUp()
        PrepareMeasure(1.5)
        WriteNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleGUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleAUp, nne)
        DrawNotesUp()
        WriteHalfNote(0, 4, bbNoteName.TrebleC, pnat)

        PrepareMeasure(2)
        WriteQuarterNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteQuarterNote(2, 4, bbNoteName.TrebleBUp, nne)
        WriteHalfNote(0, 4, bbNoteName.TrebleEUp, pnat)

        PrepareMeasure(2.5)
        WriteHalfNoteUp(0, 4, bbNoteName.TrebleBUp, nne)
        WriteHalfNote(0, 4, bbNoteName.TrebleD, pnat)

        PrepareMeasure(3)
        WriteSymbol(0, 4, 8, "h")
        WriteNote(1, 4, bbNoteName.TrebleAUp, flt)
        WriteNote(2, 4, bbNoteName.TrebleBUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleCUp, nne)
        DrawNotesUp()
        WriteHalfNote(0, 4, bbNoteName.TrebleC, pnat)

        PrepareMeasure(3.5)
        WriteNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleGUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleAUp, nne)
        DrawNotesUp()
        WriteHalfNote(0, 4, bbNoteName.TrebleD, pnat)

        PrepareMeasure(4)
        WriteQuarterNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteQuarterNote(2, 4, bbNoteName.TrebleBUp, nne)
        WriteHalfNote(0, 4, bbNoteName.TrebleEUp, pnat)

        PrepareMeasure(4.5)
        WriteHalfNoteUp(0, 4, bbNoteName.TrebleBUp, nne)
        WriteHalfNote(0, 4, bbNoteName.TrebleD, pnat)

        PrepareMeasure(5)
        WriteSymbol(0, 4, 8, "h")
        WriteNote(1, 4, bbNoteName.TrebleAUp, flt)
        WriteNote(2, 4, bbNoteName.TrebleBUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleEUp + bbNoteName.Octave, flt)
        DrawNotesUp()
        WriteHalfNote(0, 4, bbNoteName.TrebleC, pnat)

        PrepareMeasure(5.5)
        WriteNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleEUp + bbNoteName.Octave, nne)
        DrawNotesUp()

        PrepareMeasure(6)
        WriteNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleEUp + bbNoteName.Octave, nne)
        DrawNotesUp()

        PrepareMeasure(6.5)
        WriteNote(0, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleEUp + bbNoteName.Octave, nne)
        DrawNotesUp()

    End Sub

    Public Sub DrawRightBikeInnerGear(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        'Instead of calculating, just use old width from the outer gear
        'assuming that the function was called before this one
        l = Staff.StaffLength() / 12
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'Ostinato
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.2, 4, bbNoteName.TrebleG, "d")
        WriteHalfNote(3, 4, bbNoteName.TrebleEUp, pnat)
        PrepareMeasure(1.5)
        WriteHalfNote(3, 4, bbNoteName.TrebleFUp, pnat)
        PrepareMeasure(2)
        WriteHalfNote(3, 4, bbNoteName.TrebleGUp, pnat)
        PrepareMeasure(2.5)
        WriteHalfNote(3, 4, bbNoteName.TrebleAUp, pnat)
        WriteSymbol(4, 4, -10, "j")
        DrawCresc(True, 1, 3, 4, 2.5, 3, 4, -10, 4)
    End Sub
End Module
