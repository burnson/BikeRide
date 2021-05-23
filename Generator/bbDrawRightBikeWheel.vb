Public Module bbDrawRightBikeWheel
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
        n = Number * 2 - 1
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

    Public Sub DrawRightBikeWheel(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 40
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'SECTION A
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.4, 7, bbNoteName.TrebleG, "d")
        WriteEighthRest(1.3, 7)
        WriteSymbol(1.4, 7, -7, "k")
        WriteNote(2, 7, bbNoteName.TrebleGUp, shp)
        WriteNote(3.25, 7, bbNoteName.TrebleFUp, shp)
        WriteNote(4.5, 7, bbNoteName.TrebleB, pnat)
        WriteNote(5.75, 7, bbNoteName.TrebleA, pnat)
        DrawNotesRhythm2(12, 8)
        PrepareMeasure(1.5)
        WriteNote(0, 4, bbNoteName.TrebleFUp, nat)
        WriteNote(1, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(2, 4, bbNoteName.TrebleA, pnat)
        WriteNote(3, 4, bbNoteName.TrebleG, pnat)
        DrawNotes(12, 8)

        PrepareMeasure(2)
        WriteNote(0, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(1, 4, bbNoteName.TrebleD, flt)
        WriteNote(2, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotes(12, 8)
        PrepareMeasure(2.5)
        WriteNote(0, 4, bbNoteName.TrebleD, nat)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(3)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleC, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleB, flt)
        WriteNote(2.13, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotesRhythm2(8, 10)
        PrepareMeasure(3.5)
        WriteNote(0, 4, bbNoteName.TrebleD, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(4)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleC, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleB, flt)
        WriteNote(2.13, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotesRhythm2(8, 10)
        PrepareMeasure(4.5)
        WriteNote(0, 4, bbNoteName.TrebleD, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pshp)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes(8, 12)

        PrepareMeasure(5)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleFUp, shp)
        WriteNote(1.25, 4, bbNoteName.TrebleC, nne)
        WriteNote(2.13, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleA, pflt)
        DrawNotesRhythm2(12, 8)
        PrepareMeasure(5.5)
        WriteNote(0, 4, bbNoteName.TrebleFUp, shp)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleB, pnat)
        WriteNote(3, 4, bbNoteName.TrebleA, pnat)
        DrawNotes(12, 8)

        'SECTION B
        '------------------------
        PrepareMeasure(6)
        WriteNote(0, 4, bbNoteName.TrebleGUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleD, pnat)
        WriteNote(3, 4, bbNoteName.TrebleC, pnat)
        DrawNotes(12, 8)
        PrepareMeasure(6.5)
        WriteNote(0, 4, bbNoteName.TrebleEUp, nat)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleB, pnat)
        WriteNote(3, 4, bbNoteName.TrebleA, pnat)
        DrawNotes()

        PrepareMeasure(7)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleD, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleB, nne)
        WriteNote(2.13, 4, bbNoteName.TrebleA, pnat)
        WriteNote(3, 4, bbNoteName.TrebleG, pnat)
        DrawNotesRhythm2()
        PrepareMeasure(7.5)
        WriteNote(0, 4, bbNoteName.TrebleC, nne)
        WriteNote(1, 4, bbNoteName.TrebleA, nne)
        WriteNote(2, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleF, pnat)
        DrawNotes()

        PrepareMeasure(8)
        WriteNote(0, 4, bbNoteName.TrebleB, nne)
        WriteNote(1, 4, bbNoteName.TrebleG, nne)
        WriteNote(2, 4, bbNoteName.TrebleF, pnat)
        WriteNote(3, 4, bbNoteName.TrebleE, pnat)
        DrawNotes()
        PrepareMeasure(8.5)
        WriteNote(0, 4, bbNoteName.TrebleA, nat)
        WriteNote(1, 4, bbNoteName.TrebleF, nne)
        WriteNote(2, 4, bbNoteName.TrebleG, pnat)
        WriteNote(3, 4, bbNoteName.TrebleA, pnat)
        DrawNotes()

        PrepareMeasure(9)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleD, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleB, nne)
        WriteNote(2.13, 4, bbNoteName.TrebleA, pnat)
        WriteNote(3, 4, bbNoteName.TrebleG, pnat)
        DrawNotesRhythm2()
        PrepareMeasure(9.5)
        WriteNote(0, 4, bbNoteName.TrebleEUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleB, pnat)
        WriteNote(3, 4, bbNoteName.TrebleA, pnat)
        DrawNotes()

        PrepareMeasure(10)
        WriteNote(0, 4, bbNoteName.TrebleFUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleD, nne)
        WriteNote(2, 4, bbNoteName.TrebleC, pnat)
        WriteNote(3, 4, bbNoteName.TrebleB, pnat)
        DrawNotes()
        PrepareMeasure(10.5)
        WriteNote(0, 4, bbNoteName.TrebleGUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleEUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleD, pnat)
        WriteNote(3, 4, bbNoteName.TrebleC, pnat)
        DrawNotes()

        'Section C
        '-------------------
        PrepareMeasure(11)
        WriteSymbol(0, 4, 0, "h")
        WriteNote(1, 4, bbNoteName.TrebleG, pnat)
        WriteNote(2, 4, bbNoteName.TrebleEUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleC, shp)
        DrawNotes()
        PrepareMeasure(11.5)
        WriteNote(0, 4, bbNoteName.TrebleG, pnat)
        WriteNote(1, 4, bbNoteName.TrebleB, nne)
        WriteNote(2, 4, bbNoteName.TrebleD, nne)
        WriteSymbol(3, 4, 0, "h")
        DrawNotes()


        PrepareMeasure(12)
        WriteSymbol(0, 4, 0, "h")
        WriteNote(1, 4, bbNoteName.TrebleB, pnat)
        WriteNote(2, 4, bbNoteName.TrebleEUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleAUp, nne)
        DrawNotes()
        PrepareMeasure(12.5)
        WriteNote(0, 4, bbNoteName.TrebleFUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, pnat)
        WriteNote(2, 4, bbNoteName.TrebleGUp, shp)
        WriteNote(3, 4, bbNoteName.TrebleBUp, nne)
        DrawNotes()

        PrepareMeasure(13)
        WriteNote(0, 4, bbNoteName.TrebleB, pnat)
        WriteNote(1, 4, bbNoteName.TrebleFUp, shp)
        WriteNote(2, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleGUp, shp)
        DrawNotes()
        PrepareMeasure(13.5)
        WriteNote(0, 4, bbNoteName.TrebleEUp + bbNoteName.Octave, nne)
        WriteNote(1, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleCUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleGUp, pnat)
        DrawNotes()

        PrepareMeasure(14)
        WriteSymbol(0, 4, 0, "h")
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleA, pnat)
        DrawNotes()
        PrepareMeasure(14.5)
        WriteNote(0, 4, bbNoteName.TrebleFUp, shp)
        WriteNote(1, 4, bbNoteName.TrebleC, shp)
        WriteNote(2, 4, bbNoteName.TrebleD, shp)
        WriteNote(3, 4, bbNoteName.TrebleB, pnat)
        DrawNotes()

        PrepareMeasure(15)
        WriteNote(0, 4, bbNoteName.TrebleGUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nat)
        WriteNote(2, 4, bbNoteName.TrebleEUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleA, pnat)
        DrawNotes()
        PrepareMeasure(15.5)
        WriteNote(0, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(2, 4, bbNoteName.TrebleFUp, nat)
        WriteNote(3, 4, bbNoteName.TrebleB, pnat)
        DrawNotes()

        'SECTION D
        '---------------------
        PrepareMeasure(16)
        WriteNote(0, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleFUp, shp)
        WriteNote(2, 4, bbNoteName.TrebleEUp, pnat)
        WriteNote(3, 4, bbNoteName.TrebleD, pnat)
        DrawNotes()
        PrepareMeasure(16.5)
        WriteNote(0, 4, bbNoteName.TrebleGUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(2, 4, bbNoteName.TrebleEUp, pnat)
        WriteNote(3, 4, bbNoteName.TrebleD, pnat)
        DrawNotes()

        DrawCresc(True, 19, 0, 4, 20.5, 3, 4, -15, 4)
        PrepareMeasure(17)

        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleBUp, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleGUp, shp)
        WriteNote(2.13, 4, bbNoteName.TrebleFUp, pnat)
        WriteNote(3, 4, bbNoteName.TrebleEUp, pnat)
        DrawNotesRhythm2()
        PrepareMeasure(17.5)
        WriteNote(0, 4, bbNoteName.TrebleCUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleGUp, shp)
        WriteNote(2, 4, bbNoteName.TrebleAUp, pnat)
        WriteNote(3, 4, bbNoteName.TrebleGUp, pnat)
        DrawNotes()

        PrepareMeasure(18)
        WriteNote(0, 4, bbNoteName.TrebleEUp + bbNoteName.Octave, flt)
        WriteNote(1, 4, bbNoteName.TrebleCUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleBUp, pnat)
        WriteNote(3, 4, bbNoteName.TrebleAUp, pnat)
        DrawNotes()
        PrepareMeasure(18.5)
        WriteNote(0, 4, bbNoteName.TrebleGUp, nne)
        WriteNote(1, 4, bbNoteName.TrebleEUp, nne)
        WriteNote(2, 4, bbNoteName.TrebleD, pnat)
        WriteNote(3, 4, bbNoteName.TrebleC, pnat)
        DrawNotes()

        PrepareMeasure(19)
        WriteEighthRest(0, 4)
        WriteNote(0.5, 4, bbNoteName.TrebleB, nne)
        WriteNote(1.25, 4, bbNoteName.TrebleG, nne)
        WriteNote(2.13, 4, bbNoteName.TrebleDDown, pnat)
        WriteNote(3, 4, bbNoteName.TrebleCDown, pnat)
        DrawNotesRhythm2()
        PrepareMeasure(19.5)
        WriteNote(0, 4, bbNoteName.TrebleB, nne)
        WriteNote(1, 4, bbNoteName.TrebleG, nne)
        WriteNote(2, 4, bbNoteName.TrebleCDown, pnat)
        WriteNote(3, 4, bbNoteName.TrebleBDown, pnat)
        DrawNotes()

        PrepareMeasure(20)
        WriteNote(0, 4, bbNoteName.TrebleB, nne)
        WriteNote(1, 4, bbNoteName.TrebleG, nne)
        WriteNote(2, 4, bbNoteName.TrebleBDown, pnat)
        WriteNote(3, 4, bbNoteName.TrebleADown, pnat)
        DrawNotes()
        PrepareMeasure(20.5)
        WriteNote(0, 4, bbNoteName.TrebleB, nne)
        WriteNote(1, 4, bbNoteName.TrebleG, nne)
        WriteNote(2, 4, bbNoteName.TrebleADown, pnat)
        WriteNote(3, 4, bbNoteName.TrebleGDown, pnat)
        DrawNotes()
    End Sub
End Module
