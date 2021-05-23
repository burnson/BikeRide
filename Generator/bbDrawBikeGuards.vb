Module bbDrawBikeGuards
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

    Private Sub DrawCrescLeft(ByVal IsCresc As Boolean, ByVal LeftMeasure As Double, ByVal relativeXLeft As Double, ByVal relativeXLeftWidth As Double, ByVal RightMeasure As Double, ByVal relativeXRight As Double, ByVal relativeXRightWidth As Double, ByVal spacesY As Double, ByVal apex As Double)
        PrepareMeasureLeft(LeftMeasure)
        Dim x1 As Double = GetMeasuredX(LeftMeasure, relativeXLeft, relativeXLeftWidth)
        PrepareMeasureLeft(RightMeasure)
        Dim x2 As Double = GetMeasuredX(RightMeasure, relativeXRight, relativeXRightWidth)
        DrawHairpin(Staff, IsCresc, x1, spacesY, x2, spacesY, apex, Globals.ratioStaffLineToSpaceHeight, Globals.numberBeamInterpolations, New bbColor(0, 0, 0))
    End Sub


    Private Sub DrawCrescRight(ByVal IsCresc As Boolean, ByVal LeftMeasure As Double, ByVal relativeXLeft As Double, ByVal relativeXLeftWidth As Double, ByVal RightMeasure As Double, ByVal relativeXRight As Double, ByVal relativeXRightWidth As Double, ByVal spacesY As Double, ByVal apex As Double)
        PrepareMeasureRight(LeftMeasure)
        Dim x1 As Double = GetMeasuredX(LeftMeasure, relativeXLeft, relativeXLeftWidth)
        PrepareMeasureRight(RightMeasure)
        Dim x2 As Double = GetMeasuredX(RightMeasure, relativeXRight, relativeXRightWidth)
        DrawHairpin(Staff, IsCresc, x1, spacesY, x2, spacesY, apex, Globals.ratioStaffLineToSpaceHeight, Globals.numberBeamInterpolations, New bbColor(0, 0, 0))
    End Sub

    Private Sub PrepareMeasure(ByVal Number As Double)
        n = Number * 2 - 1
        m = New Collection
        o = l * (n - 1) + l * Globals.percentageMeasureLeftMargin
    End Sub

    Private Sub PrepareMeasureLeft(ByVal Number As Double)
        n = Number + 12
        m = New Collection
        o = l * (n - 1) + l * Globals.percentageMeasureLeftMargin
    End Sub

    Private Sub PrepareMeasureRight(ByVal Number As Double)
        n = 13 - Number
        m = New Collection
        o = l * (n - 1) + l * Globals.percentageMeasureLeftMargin
    End Sub

    Private Function GetMeasuredX(ByVal Measure As Double, ByVal relativeX As Double, ByVal relativeWidth As Double) As Double
        Return (relativeX + 0.7) / (relativeWidth + 0.2) * w + o
    End Function

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

    Private Sub WriteHalfNote(ByVal relativeX As Double, ByVal relativeWidth As Double, ByVal NoteName As bbNoteName, ByVal NoteAccidental As bbAccidentalType)
        Dim x As Double = (relativeX + 0.5) / relativeWidth * w + o
        DrawNoteWithStem(Staff, x, NoteName, False, False, StemThickness, 12, c, NoteAccidental, NotationFont)
    End Sub

    Private Sub WriteEighthRest(ByVal relativeX As Double, ByVal relativeWidth As Double)
        WriteSymbol(relativeX, relativeWidth, 0, "h")
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

    Private Sub DrawNotesRhythm3(ByVal LeftHeight As Double, ByVal RightHeight As Double)
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m, LeftHeight, RightHeight)
        DrawRhythm3(Staff, m, NotationFont)
    End Sub

    Private Sub DrawNotesRhythm3()
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m)
        DrawRhythm3(Staff, m, NotationFont)
    End Sub

    Private Sub DrawNotesRhythm4(ByVal LeftHeight As Double, ByVal RightHeight As Double)
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m, LeftHeight, RightHeight)
        DrawRhythm4(Staff, m, NotationFont)
    End Sub

    Private Sub DrawNotesRhythm4()
        SetBasicNoteGroupStemDirections(m)
        SetBasicNoteGroupStemHeights(m)
        DrawRhythm4(Staff, m, NotationFont)
    End Sub

    Public Sub DrawLeftBikeGuard(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 24
        w = l - (l * Globals.percentageMeasureLeftMargin)

        PrepareMeasureLeft(1) '1
        WriteSymbol(0.0, 3.8, bbNoteName.BassF, "e")
        WriteSymbol(0.8, 3.8, -16, "j")
        WriteNote(0.8, 3.8, bbNoteName.BassCDown, nne)
        WriteNote(1.8, 3.8, bbNoteName.BassG, nne)
        WriteNote(2.8, 3.8, bbNoteName.BassE, nne)
        DrawNotes(17, 8)
        PrepareMeasureLeft(2) '2
        WriteNote(0, 3, bbNoteName.BassDDown, flt)
        WriteNote(1, 3, bbNoteName.BassA, flt)
        WriteNote(2, 3, bbNoteName.BassF, nne)
        DrawNotes(17, 8)

        PrepareMeasureLeft(3) '3
        WriteNote(0, 3, bbNoteName.BassDDown, nat)
        WriteNote(1, 3, bbNoteName.BassA, nat)
        WriteNote(2, 3, bbNoteName.BassF, shp)
        DrawNotes(17, 8)
        PrepareMeasureLeft(4) '4
        WriteNote(0, 3, bbNoteName.BassEDown, flt)
        WriteNote(1, 3, bbNoteName.BassB, flt)
        WriteNote(2, 3, bbNoteName.BassGUp, nne)
        DrawNotes(17, 8)

        PrepareMeasureLeft(5) '5
        WriteNote(0, 3, bbNoteName.BassEDown, nat)
        WriteNote(1, 3, bbNoteName.BassB, nat)
        WriteNote(2, 3, bbNoteName.BassGUp, shp)
        DrawNotes(17, 8)
        PrepareMeasureLeft(6) '6
        WriteNote(0, 3, bbNoteName.BassFDown, nat)
        WriteNote(1, 3, bbNoteName.BassC, nat)
        WriteNote(2, 3, bbNoteName.BassAUp, nne)
        DrawNotes(17, 8)

        PrepareMeasureLeft(7) '7
        WriteNote(0, 3, bbNoteName.BassG, flt)
        WriteNote(1, 3, bbNoteName.BassD, flt)
        WriteNote(2, 3, bbNoteName.BassBUp, flt)
        DrawNotes(8, 17)
        PrepareMeasureLeft(8) '8
        WriteNote(0, 3, bbNoteName.BassG, nat)
        WriteNote(1, 3, bbNoteName.BassD, nat)
        WriteNote(2, 3, bbNoteName.BassBUp, nat)
        DrawNotes(8, 17)

        PrepareMeasureLeft(9) '9
        WriteNote(0, 3, bbNoteName.BassA, flt)
        WriteNote(1, 3, bbNoteName.BassE, flt)
        WriteNote(2, 3, bbNoteName.BassCUp, nne)
        DrawNotes(8, 17)

        PrepareMeasureLeft(10) '10
        WriteNote(0, 3, bbNoteName.BassA, nat)
        WriteNote(1, 3, bbNoteName.BassE, nat)
        WriteNote(2, 3, bbNoteName.BassCUp, shp)
        DrawNotes(8, 17)

        PrepareMeasureLeft(11) '11
        WriteEighthRest(0, 2)
        WriteEighthRest(1, 2)

        PrepareMeasureLeft(12) '12-12
        WriteSymbol(0.0, 3.8, bbNoteName.TrebleG, "d") 'treble clef + 8
        WriteHalfNote(1, 3.8, bbNoteName.TrebleGUp + bbNoteName.Octave, shp)
        WriteSymbol(1, 3.8, -16, "o")
        DrawCrescLeft(True, 1, 1.8, 3.8, 12, 0, 3.8, -16, 4)
    End Sub

    Public Sub DrawRightBikeGuard(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        'Instead of calculating, just use old width from the outer gear
        'assuming that the function was called before this one
        l = Staff.StaffLength() / 12
        w = l - (l * Globals.percentageMeasureLeftMargin)

        PrepareMeasureRight(1)
        WriteNote(0, 6, bbNoteName.TrebleAUp, pnat)
        WriteNote(1, 6, bbNoteName.TrebleDUp, shp)
        WriteNote(2, 6, bbNoteName.TrebleFUp + bbNoteName.Octave, shp)
        WriteNote(3, 6, bbNoteName.TrebleGUp + bbNoteName.Octave, shp)
        WriteNote(4, 6, bbNoteName.TrebleAUp + bbNoteName.Octave, shp)
        WriteSymbol(5, 6, bbNoteName.TrebleG, "d") 'treble clef + 8va
        WriteSymbol(4, 6, -16, "j")
        DrawNotesRhythm3()

        PrepareMeasureRight(2)
        WriteNote(0, 5, bbNoteName.TrebleGUp, pnat)
        WriteNote(1, 5, bbNoteName.TrebleCUp, shp)
        WriteNote(2, 5, bbNoteName.TrebleDUp, shp)
        WriteNote(3, 5, bbNoteName.TrebleFUp + bbNoteName.Octave, shp)
        WriteNote(4, 5, bbNoteName.TrebleGUp + bbNoteName.Octave, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(3)
        WriteNote(0, 5, bbNoteName.TrebleFUp, pnat)
        WriteNote(1, 5, bbNoteName.TrebleAUp, shp)
        WriteNote(2, 5, bbNoteName.TrebleCUp, shp)
        WriteNote(3, 5, bbNoteName.TrebleDUp, shp)
        WriteNote(4, 5, bbNoteName.TrebleFUp + bbNoteName.Octave, shp)
        DrawNotesRhythm3()


        PrepareMeasureRight(4)
        WriteNote(0, 5, bbNoteName.TrebleD, pnat)
        WriteNote(1, 5, bbNoteName.TrebleGUp, shp)
        WriteNote(2, 5, bbNoteName.TrebleAUp, shp)
        WriteNote(3, 5, bbNoteName.TrebleCUp, shp)
        WriteNote(4, 5, bbNoteName.TrebleDUp, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(5)
        WriteNote(0, 5, bbNoteName.TrebleC, pnat)
        WriteNote(1, 5, bbNoteName.TrebleFUp, shp)
        WriteNote(2, 5, bbNoteName.TrebleGUp, shp)
        WriteNote(3, 5, bbNoteName.TrebleAUp, shp)
        WriteNote(4, 5, bbNoteName.TrebleCUp, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(6)
        WriteNote(0, 5, bbNoteName.TrebleA, pnat)
        WriteNote(1, 5, bbNoteName.TrebleD, shp)
        WriteNote(2, 5, bbNoteName.TrebleFUp, shp)
        WriteNote(3, 5, bbNoteName.TrebleGUp, shp)
        WriteNote(4, 5, bbNoteName.TrebleAUp, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(7)
        WriteNote(0, 5, bbNoteName.TrebleG, pnat)
        WriteNote(1, 5, bbNoteName.TrebleC, shp)
        WriteNote(2, 5, bbNoteName.TrebleD, shp)
        WriteNote(3, 5, bbNoteName.TrebleFUp, shp)
        WriteNote(4, 5, bbNoteName.TrebleGUp, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(8)
        WriteNote(0, 5, bbNoteName.TrebleF, pnat)
        WriteNote(1, 5, bbNoteName.TrebleA, shp)
        WriteNote(2, 5, bbNoteName.TrebleC, shp)
        WriteNote(3, 5, bbNoteName.TrebleD, shp)
        WriteNote(4, 5, bbNoteName.TrebleFUp, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(9)
        WriteNote(0, 5, bbNoteName.TrebleDDown, pnat)
        WriteNote(1, 5, bbNoteName.TrebleG, shp)
        WriteNote(2, 5, bbNoteName.TrebleA, shp)
        WriteNote(3, 5, bbNoteName.TrebleC, shp)
        WriteNote(4, 5, bbNoteName.TrebleD, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(10)
        WriteNote(0, 5, bbNoteName.TrebleF, pnat)
        WriteNote(1, 5, bbNoteName.TrebleA, shp)
        WriteNote(2, 5, bbNoteName.TrebleC, shp)
        WriteNote(3, 5, bbNoteName.TrebleD, shp)
        WriteNote(4, 5, bbNoteName.TrebleFUp, shp)
        DrawNotesRhythm3()

        PrepareMeasureRight(11)
        WriteNote(0, 4, bbNoteName.TrebleFUp, shp)
        WriteNote(1, 4, bbNoteName.TrebleD, shp)
        WriteNote(2, 4, bbNoteName.TrebleC, shp)
        WriteNote(3, 4, bbNoteName.TrebleA, shp)
        WriteSymbol(0, 4, -16, "o")
        DrawNotesRhythm3()

        DrawCrescRight(False, 11, 1, 4, 1, 3, 6, -16, 4)
    End Sub
End Module
