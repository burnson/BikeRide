Module bbDrawGroundNotes
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
        Return (relativeX + 0.7) / (relativeWidth + 0.2) * w + o
    End Function

    Private Sub PrepareMeasure(ByVal Number As Double)
        n = Number
        m = New Collection
        o = l * (n - 1) + l * Globals.percentageMeasureLeftMargin
    End Sub

    Private Sub WriteNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal noteName As bbNoteName, ByVal noteAccidental As bbAccidentalType)
        m.Add(New bbNoteProperty((relativeX + 0.7) / (relativeWidth + 0.2), noteName, noteAccidental, w, o, c, t))
    End Sub

    Private Sub WriteSymbol(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal spacesY As Double, ByVal symbol As String)
        Dim LocalSpaceHeight As Double
        Dim Orientation As bbRay = Nothing
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        Staff.FindPlottingPoint(x, spacesY, Orientation, LocalSpaceHeight, 0)
        Page.AddLineToStream(DrawString(NotationFont, LocalSpaceHeight, Orientation, symbol))
    End Sub

    Private Sub WriteEighthRest(ByVal relativeX As Double, ByVal relativeWidth As Double)
        WriteSymbol(relativeX, relativeWidth, 0, "h")
    End Sub

    Private Sub WriteQuarterNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal NoteName As bbNoteName, ByVal NoteAccidental As bbAccidentalType)
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        DrawNoteWithStem(Staff, x, NoteName, True, True, StemThickness, Globals.spacesDefaultStemHeight, c, NoteAccidental, NotationFont)
    End Sub

    Private Sub WriteEighthNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal NoteName As bbNoteName, ByVal NoteAccidental As bbAccidentalType)
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        DrawNoteWithStemAndFlag(Staff, x, NoteName, True, True, StemThickness, Globals.spacesDefaultStemHeight, c, NoteAccidental, NotationFont)
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

    Private Sub DrawNotesRhythm2(ByVal LeftHeight As Double, ByVal RightHeight As Double)
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m, LeftHeight, RightHeight)
        DrawRhythm2(Staff, m, NotationFont)
    End Sub

    Private Sub DrawNotesRhythm2()
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m)
        DrawRhythm2(Staff, m, NotationFont)
    End Sub

    Public Sub DrawLeftGroundNotes(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 12.5
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'SECTION A
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.7, 7, bbNoteName.BassF, "e")
        WriteSymbol(2, 7, -14, "j")
        WriteNote(2, 7, bbNoteName.BassB, flt)
        WriteNote(3.25, 7, bbNoteName.BassF, nne)
        WriteNote(4.5, 7, bbNoteName.BassDUp, nne)
        WriteNote(5.75, 7, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(2)
        WriteNote(0, 4, bbNoteName.BassE, flt)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(3)
        WriteNote(0, 4, bbNoteName.BassB, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(4)
        WriteNote(0, 4, bbNoteName.BassE, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(5)
        WriteNote(0, 4, bbNoteName.BassB, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(6)
        WriteNote(0, 4, bbNoteName.BassE, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(7)
        WriteNote(0, 4, bbNoteName.BassB, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(8)
        WriteNote(0, 4, bbNoteName.BassE, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(9)
        WriteNote(0, 4, bbNoteName.BassB, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(10)
        WriteNote(0, 4, bbNoteName.BassE, nne)
        WriteNote(1, 4, bbNoteName.BassF, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes()

        PrepareMeasure(11)
        WriteEighthNote(0, 4, bbNoteName.BassB, flt)
        WriteEighthRest(1, 4)
        WriteEighthRest(2, 4)
        WriteEighthRest(3, 4)

        PrepareMeasure(12)
        WriteEighthNote(0, 4, bbNoteName.BassE, flt)
        WriteEighthRest(1, 4)
        WriteEighthRest(2, 4)
        WriteEighthRest(3, 4)

        PrepareMeasure(13)
        WriteQuarterNote(0, 4, bbNoteName.BassB, flt)

        'DrawCresc(False, 11, 0, 4, 12, 4, 4, -12, 4)

        DrawCresc(False, 10, 4.8, 4.9, 12, 3, 4, -12, 4)
    End Sub

    Public Sub DrawRightGroundNotes(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 12.5
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'SECTION A
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.4, 7, bbNoteName.TrebleG, "d")
        WriteEighthRest(1.3, 7)
        WriteSymbol(2, 7, -14, "j")
        WriteNote(2, 7, bbNoteName.TrebleGUp, shp)
        WriteNote(3.25, 7, bbNoteName.TrebleFUp, shp)
        WriteNote(4.5, 7, bbNoteName.TrebleB, pnat)
        WriteNote(5.75, 7, bbNoteName.TrebleA, pnat)
        DrawNotesRhythm2(12, 8)
        PrepareMeasure(2)
        WriteNote(0, 4, bbNoteName.TrebleFUp, nat)
        WriteNote(1, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(2, 4, bbNoteName.TrebleA, pnat)
        WriteNote(3, 4, bbNoteName.TrebleG, pnat)
        DrawNotes(12, 8)

        PrepareMeasure(3)
        WriteNote(0, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(1, 4, bbNoteName.TrebleD, flt)
        WriteNote(2, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotes(12, 8)
        PrepareMeasure(4)
        WriteNote(0, 4, bbNoteName.TrebleD, nat)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(5)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleC, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleB, flt)
        WriteNote(2.13, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotesRhythm2(8, 10)
        PrepareMeasure(6)
        WriteNote(0, 4, bbNoteName.TrebleD, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(7)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleC, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleB, flt)
        WriteNote(2.13, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotesRhythm2(8, 10)
        PrepareMeasure(8)
        WriteNote(0, 4, bbNoteName.TrebleD, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(9)
        WriteNote(0, 4, bbNoteName.TrebleC, nne)
        WriteNote(1, 4, bbNoteName.TrebleB, flt)
        WriteNote(2, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotes(8, 10)

        PrepareMeasure(10)
        WriteNote(0, 4, bbNoteName.TrebleD, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(11)
        WriteSymbol(0, 4.8, bbNoteName.BassF, "e")
        WriteEighthRest(0.8, 3.8)
        WriteNote(1.8, 4.8, bbNoteName.BassF, nat)
        WriteNote(2.8, 4.8, bbNoteName.BassDUp, nne)
        WriteNote(3.8, 4.8, bbNoteName.BassF, nne)
        DrawNotes(8, 10)

        PrepareMeasure(12)
        WriteEighthRest(0, 4)
        WriteNote(1, 4, bbNoteName.BassE, pnat)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassF, nne)
        DrawNotes(8, 12)

        DrawCresc(False, 10, 4.8, 4.9, 12, 3, 4, -12, 4)
    End Sub
End Module
