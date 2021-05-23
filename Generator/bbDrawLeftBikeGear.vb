Module bbDrawLeftBikeGear
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

    Public Sub DrawLeftBikeGear(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        l = Staff.StaffLength() / 12
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'Ostinato
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.5, 4, bbNoteName.TrebleG, "d")
        WriteNote(1.5, 4, bbNoteName.TrebleADown, flt)
        WriteNote(2.5, 4, bbNoteName.TrebleE, flt)
        WriteNote(3.5, 4, bbNoteName.TrebleB, nne)
        WriteNote(4.5, 4, bbNoteName.TrebleE, nne)
        DrawNotes(14, 10)
        PrepareMeasure(1.5)
        WriteNote(1.5, 4, bbNoteName.TrebleDDown, flt)
        WriteNote(2.5, 4, bbNoteName.TrebleE, nne)
        WriteNote(3.5, 4, bbNoteName.TrebleB, nne)
        WriteNote(4.5, 4, bbNoteName.TrebleE, nne)
        DrawNotes(11, 10)
    End Sub

    Public Sub DrawLeftBikeInnerGear(ByRef refStaff As bbStaff)
        Staff = refStaff
        Staff.GetParent(Page)
        'Instead of calculating, just use old width from the outer gear
        'assuming that the function was called before this one
        'l = Staff.StaffLength() / 8 
        w = l - (l * Globals.percentageMeasureLeftMargin)

        'Ostinato
        '------------------------
        PrepareMeasure(1)
        WriteSymbol(0.5, 4, bbNoteName.BassF, "e")
        WriteNote(1.5, 4, bbNoteName.BassB, flt)
        WriteNote(2.5, 4, bbNoteName.BassF, nne)
        WriteNote(3.5, 4, bbNoteName.BassDUp, nne)
        WriteNote(4.5, 4, bbNoteName.BassF, nne)
        DrawNotes(8, 12)
        PrepareMeasure(1.5)
        WriteNote(1.5, 4, bbNoteName.BassE, flt)
        WriteNote(2.5, 4, bbNoteName.BassF, nne)
        WriteNote(3.5, 4, bbNoteName.BassDUp, nne)
        WriteNote(4.5, 4, bbNoteName.BassF, nne)
        DrawNotes(11, 12)
    End Sub
End Module
