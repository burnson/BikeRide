Public Class bbNoteProperty
    Public NoteName As bbNoteName
    Public Accidental As bbAccidentalType
    Public StemUp As Boolean
    Public StemHeight As Double
    Public unitsX As Double
    Public StemWidth As Double
    Public Color As bbColor

    Public Sub New(ByVal x As Double, ByVal Note As bbNoteName, ByVal AccidentalType As bbAccidentalType, ByVal GroupWidth As Double, ByVal GroupOffset As Double, ByVal NoteColor As bbColor, ByVal unitsStemWidth As Double)
        unitsX = x * GroupWidth + GroupOffset
        NoteName = Note
        Accidental = AccidentalType
        StemUp = True
        StemHeight = Globals.spacesDefaultStemHeight
        Color = NoteColor
        StemWidth = unitsStemWidth
    End Sub

    Public Sub New()
        unitsX = 0
        NoteName = 0
        Accidental = 0
        StemUp = True
        StemHeight = Globals.spacesDefaultStemHeight
        Color = New bbColor(0, 0, 0)
        StemWidth = 0
    End Sub
End Class
