Public Module bbDrawLeftBikeWheel
    Public Const StemThickness As Double = 0.005
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


    Public Sub DrawLeftBikeWheel(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 40
        w = l - (l * Globals.percentageMeasureLeftMargin)

        DrawCresc(True, 19, 0, 4, 20.5, 3, 4, -15, 4)

        'SECTION A
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.7, 7, bbNoteName.BassF, "e")
        WriteNote(2, 7, bbNoteName.BassF, shp)
        WriteSymbol(2, 7, -10, "k")
        WriteNote(3.25, 7, bbNoteName.BassAUp, shp)
        WriteNote(4.5, 7, bbNoteName.BassFUp, nat)
        WriteNote(5.75, 7, bbNoteName.BassAUp, nne)
        DrawNotes()
        PrepareMeasure(1.5)
        WriteNote(0, 4, bbNoteName.BassGUp, nne)
        WriteNote(1, 4, bbNoteName.BassBUp, nne)
        WriteNote(2, 4, bbNoteName.BassFUp, nne)
        WriteNote(3, 4, bbNoteName.BassBUp, nne)
        DrawNotes()
        'Exit Sub

        PrepareMeasure(2)
        WriteNote(0, 4, bbNoteName.BassE, nne)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()
        PrepareMeasure(2.5)
        WriteNote(0, 4, bbNoteName.BassGUp, shp)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        PrepareMeasure(3)
        WriteNote(0, 4, bbNoteName.BassAUp, nat)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()
        PrepareMeasure(3.5)
        WriteNote(0, 4, bbNoteName.BassGUp, shp)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        PrepareMeasure(4)
        WriteNote(0, 4, bbNoteName.BassAUp, nne)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()
        PrepareMeasure(4.5)
        WriteNote(0, 4, bbNoteName.BassGUp, shp)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        PrepareMeasure(5)
        WriteNote(0, 4, bbNoteName.BassGUp, nat)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()
        PrepareMeasure(5.5)
        WriteNote(0, 4, bbNoteName.BassF, nne)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, flt)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        'SECTION B
        '-----------------------
        PrepareMeasure(6)
        WriteNote(0, 4, bbNoteName.BassE, nat)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassGUp + bbNoteName.Octave, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()
        PrepareMeasure(6.5)
        WriteNote(0, 4, bbNoteName.BassF, nne)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, flt)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        PrepareMeasure(7)
        WriteNote(0, 4, bbNoteName.BassGUp, nne)
        WriteNote(1, 4, bbNoteName.BassBUp, nne)
        WriteNote(2, 4, bbNoteName.BassCUp, nne)
        WriteNote(3, 4, bbNoteName.BassBUp, nne)
        DrawNotes()
        PrepareMeasure(7.5)
        WriteNote(0, 4, bbNoteName.BassAUp, nne)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassDUp, nne)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        PrepareMeasure(8)
        WriteNote(0, 4, bbNoteName.BassBUp, nne)
        WriteNote(1, 4, bbNoteName.BassDUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, nat)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()
        PrepareMeasure(8.5)
        WriteNote(0, 4, bbNoteName.BassGUp, nne)
        WriteNote(1, 4, bbNoteName.BassDUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, nne)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()

        PrepareMeasure(9)
        WriteNote(0, 4, bbNoteName.BassCUp, nne)
        WriteNote(1, 4, bbNoteName.BassDUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, nne)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()
        PrepareMeasure(9.5)
        WriteNote(0, 4, bbNoteName.BassBUp, nne)
        WriteNote(1, 4, bbNoteName.BassDUp, nne)
        WriteNote(2, 4, bbNoteName.BassEUp, nne)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()

        PrepareMeasure(10)
        WriteNote(0, 4, bbNoteName.BassAUp, nne)
        WriteNote(1, 4, bbNoteName.BassDUp, nne)
        WriteNote(2, 4, bbNoteName.BassGUp + bbNoteName.Octave, nne)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()
        PrepareMeasure(10.5)
        WriteNote(0, 4, bbNoteName.BassGUp, shp)
        WriteNote(1, 4, bbNoteName.BassDUp, nne)
        WriteNote(2, 4, bbNoteName.BassFUp, shp)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()

        'SECTION C
        '------------------
        PrepareMeasure(11)
        WriteNote(0, 4, bbNoteName.BassGUp, nne)
        WriteNote(1, 4, bbNoteName.BassBUp, nne)
        WriteNote(2, 4, bbNoteName.BassFUp, nne)
        WriteNote(3, 4, bbNoteName.BassDUp, nne)
        DrawNotes()
        PrepareMeasure(11.5)
        WriteNote(0, 4, bbNoteName.BassF, nne)
        WriteNote(1, 4, bbNoteName.BassDUp, shp)
        WriteNote(2, 4, bbNoteName.BassGUp + bbNoteName.Octave, nne)
        WriteNote(3, 4, bbNoteName.BassEUp, nne)
        DrawNotes()

        PrepareMeasure(12)
        WriteNote(0, 4, bbNoteName.BassE, nne)
        WriteNote(1, 4, bbNoteName.BassGUp + bbNoteName.Octave, flt)
        WriteNote(2, 4, bbNoteName.BassBUp + bbNoteName.Octave, nne)
        WriteNote(3, 4, bbNoteName.BassFUp, nne)
        DrawNotes()
        PrepareMeasure(12.5)
        WriteNote(0, 4, bbNoteName.BassE, flt)
        WriteNote(1, 4, bbNoteName.BassAUp + bbNoteName.Octave, flt)
        WriteNote(2, 4, bbNoteName.BassBUp + bbNoteName.Octave, nne)
        WriteNote(3, 4, bbNoteName.BassGUp + bbNoteName.Octave, nne)
        DrawNotes()

        PrepareMeasure(13)
        WriteNote(0, 4, bbNoteName.BassD, nne)
        WriteNote(1, 4, bbNoteName.BassGUp + bbNoteName.Octave, shp)
        WriteNote(2, 4, bbNoteName.BassCUp + bbNoteName.Octave, nne)
        WriteNote(3, 4, bbNoteName.BassAUp + bbNoteName.Octave, nne)
        DrawNotes()
        PrepareMeasure(13.5)
        WriteNote(0, 4, bbNoteName.BassFUp + bbNoteName.Octave, shp)
        WriteNote(1, 4, bbNoteName.BassFUp + bbNoteName.Octave, nat)
        WriteNote(2, 4, bbNoteName.BassEUp + bbNoteName.Octave, nne)
        WriteEighthRest(3, 4)
        DrawNotes(16, 16)

        PrepareMeasure(14)
        WriteNote(0, 4, bbNoteName.BassAUp, flt)
        WriteNote(1, 4, bbNoteName.BassCUp, nne)
        WriteNote(2, 4, bbNoteName.BassGUp + bbNoteName.Octave, flt)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()
        PrepareMeasure(14.5)
        WriteNote(0, 4, bbNoteName.BassAUp, nat)
        WriteNote(1, 4, bbNoteName.BassCUp, shp)
        WriteNote(2, 4, bbNoteName.BassGUp + bbNoteName.Octave, nat)
        WriteNote(3, 4, bbNoteName.BassCUp, nne)
        DrawNotes()

        PrepareMeasure(15)
        WriteNote(0, 4, bbNoteName.BassBUp, flt)
        WriteNote(1, 4, bbNoteName.BassEUp, flt)
        WriteNote(2, 4, bbNoteName.BassAUp + bbNoteName.Octave, flt)
        WriteNote(3, 4, bbNoteName.BassEUp, nne)
        DrawNotes()
        PrepareMeasure(15.5)
        WriteNote(0, 4, bbNoteName.BassGUp, flt)
        WriteNote(1, 4, bbNoteName.BassEUp, nat)
        WriteNote(2, 4, bbNoteName.BassBUp + bbNoteName.Octave, flt)
        WriteNote(3, 4, bbNoteName.BassEUp, nne)
        DrawNotes()

        'SECTION D
        '---------------------------
        PrepareMeasure(16)
        WriteSymbol(0.7, 7, bbNoteName.TrebleG, "d")
        WriteNote(2, 7, bbNoteName.TrebleBDown, flt)
        WriteNote(3.25, 7, bbNoteName.TrebleDDown, nne)
        WriteNote(4.5, 7, bbNoteName.TrebleA, flt)
        WriteNote(5.75, 7, bbNoteName.TrebleDDown, nne)
        DrawNotes(12, 12)
        PrepareMeasure(16.5)
        WriteNote(0, 4, bbNoteName.TrebleBDown, nat)
        WriteNote(1, 4, bbNoteName.TrebleDDown, shp)
        WriteNote(2, 4, bbNoteName.TrebleA, nat)
        WriteNote(3, 4, bbNoteName.TrebleDDown, nne)
        DrawNotes(12, 12)

        PrepareMeasure(17)
        WriteNote(0, 4, bbNoteName.TrebleCDown, nne)
        WriteNote(1, 4, bbNoteName.TrebleA, nne)
        WriteNote(2, 4, bbNoteName.TrebleC, shp)
        WriteNote(3, 4, bbNoteName.TrebleA, nne)
        DrawNotes(12, 8)
        PrepareMeasure(17.5)
        WriteNote(0, 4, bbNoteName.TrebleCDown, shp)
        WriteNote(1, 4, bbNoteName.TrebleB, flt)
        WriteNote(2, 4, bbNoteName.TrebleD, nne)
        WriteNote(3, 4, bbNoteName.TrebleB, nne)
        DrawNotes(12, 8)

        PrepareMeasure(18)
        WriteNote(0, 4, bbNoteName.TrebleDDown, nne)
        WriteNote(1, 4, bbNoteName.TrebleC, nat)
        WriteNote(2, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleC, nne)
        DrawNotes(6, 12)
        PrepareMeasure(18.5)
        WriteNote(0, 4, bbNoteName.TrebleE, nat)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleEUp, flt)
        WriteNote(3, 4, bbNoteName.TrebleD, flt)
        DrawNotes(6, 12)

        PrepareMeasure(19)
        WriteNote(0, 4, bbNoteName.TrebleF, shp)
        WriteNote(1, 4, bbNoteName.TrebleC, nne)
        WriteNote(2, 4, bbNoteName.TrebleGUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleC, nne)
        DrawNotes()
        PrepareMeasure(19.5)
        WriteNote(0, 4, bbNoteName.TrebleF, shp)
        WriteNote(1, 4, bbNoteName.TrebleC, shp)
        WriteNote(2, 4, bbNoteName.TrebleGUp, shp)
        WriteNote(3, 4, bbNoteName.TrebleC, nne)
        DrawNotes()

        PrepareMeasure(20)
        WriteNote(0, 4, bbNoteName.TrebleF, shp)
        WriteNote(1, 4, bbNoteName.TrebleD, nne)
        WriteNote(2, 4, bbNoteName.TrebleAUp, nne)
        WriteNote(3, 4, bbNoteName.TrebleD, nne)
        DrawNotes()
        PrepareMeasure(20.5)
        WriteNote(0, 4, bbNoteName.TrebleF, shp)
        WriteNote(1, 4, bbNoteName.TrebleD, shp)
        WriteNote(2, 4, bbNoteName.TrebleAUp, shp)
        WriteNote(3, 4, bbNoteName.TrebleD, nne)
        DrawNotes()
    End Sub
End Module
