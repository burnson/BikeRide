Public Module bbGlobalConstants
    Public Globals As bbGlobal
    Private States As New Collection

    Public Structure bbGlobal
        Public ratioBarLineToStaffLineThickness As Double
        Public ratioLedgerLineToSpaceHeight As Double
        Public ratioNoteWidthToNoteHeight As Double
        Public degreesNoteRotation As Double
        Public radiansNoteRotation As Double
        Public ratioAccidentalDisplacementToSpaceHeight As Double
        Public spacesDefaultStemHeight As Double
        Public numberBeamInterpolations As Double
        Public ratioBeamThicknessToSpaceHeight As Double
        Public spacesBeamSeparation As Double
        Public percentageMeasureLeftMargin As Double
        Public ratioStaffLineToSpaceHeight As Double


        Public Sub SetDefaults()
            ratioBarLineToStaffLineThickness = 1.5
            ratioLedgerLineToSpaceHeight = 2.0
            ratioNoteWidthToNoteHeight = 1.25
            degreesNoteRotation = 30
            radiansNoteRotation = (Globals.degreesNoteRotation / 360.0) * Math.PI * 2.0
            ratioAccidentalDisplacementToSpaceHeight = 0.8
            spacesDefaultStemHeight = 8
            numberBeamInterpolations = 40
            ratioBeamThicknessToSpaceHeight = 0.7
            spacesBeamSeparation = 0.6 + ratioBeamThicknessToSpaceHeight * 2.0
            percentageMeasureLeftMargin = 0
            ratioStaffLineToSpaceHeight = 0.23
        End Sub

        Public Sub SaveState()
            SaveGlobalState()
        End Sub

        Public Sub RevertState()
            RevertGlobalState()
        End Sub
    End Structure

    Public Enum bbAccidentalType As Integer
        None = 0
        Natural = 1
        Sharp = 2
        Flat = 3
        NaturalPairUp = 4
        SharpPairUp = 5
        FlatPairUp = 6
    End Enum
    Public Enum bbNoteName As Integer
        TrebleEDown = -11
        TrebleFDown = -10
        TrebleGDown = -9
        TrebleADown = -8
        TrebleBDown = -7
        TrebleCDown = -6
        TrebleDDown = -5
        TrebleE = -4
        TrebleF = -3
        TrebleG = -2
        TrebleA = -1
        TrebleB = 0
        TrebleC = 1
        TrebleD = 2
        TrebleEUp = 3
        TrebleFUp = 4
        TrebleGUp = 5
        TrebleAUp = 6
        TrebleBUp = 7
        TrebleCUp = 8
        TrebleDUp = 9

        BassGDown = -11
        BassADown = -10
        BassBDown = -9
        BassCDown = -8
        BassDDown = -7
        BassEDown = -6
        BassFDown = -5
        BassG = -4
        BassA = -3
        BassB = -2
        BassC = -1
        BassD = 0
        BassE = 1
        BassF = 2
        BassGUp = 3
        BassAUp = 4
        BassBUp = 5
        BassCUp = 6
        BassDUp = 7
        BassEUp = 8
        BassFUp = 9

        Octave = 7
    End Enum

    Private Sub SaveGlobalState()
        Dim CurrentState As New bbGlobal
        CurrentState = Globals
        States.Add(CurrentState)
    End Sub

    Private Sub RevertGlobalState()
        Globals = States(States.Count())
        States.Remove(States.Count())
    End Sub
End Module
