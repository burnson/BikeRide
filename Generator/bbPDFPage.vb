Public Class bbPDFPage
    Private WidthInches As Double
    Private HeightInches As Double
    Private PageStream As New System.Text.StringBuilder(1024 * 16)

    Private BytePageStream() As Byte
    Private BytePageStreamVirtualLength As Integer
    Private BytePageStreamActualLength As Integer

    Public Sub New(ByVal WidthInInches As Double, ByVal HeightInInches As Double)
        WidthInches = WidthInInches
        HeightInches = HeightInInches

        AddLineToStream(CTMScale(72, 72))
    End Sub

    Public Sub AddLineToStream(ByRef Str As String)
        AddLineToStreamTime.StartClock()
        Dim x As Integer = Len(Str)
        Dim y As Integer = Len(TokenSeperator)
        Do While PageStream.Length() + x + y >= PageStream.Capacity()
            PageStream.Capacity *= 2
        Loop
        PageStream.Append(Str)
        PageStream.Append(TokenSeperator())
        AddLineToStreamTime.StopClock()
    End Sub

    Public ReadOnly Property Width() As Double
        Get
            Return WidthInches
        End Get
    End Property

    Public ReadOnly Property Height() As Double
        Get
            Return HeightInches
        End Get
    End Property

    Public ReadOnly Property StreamLength() As Integer
        Get
            Return PageStream.Length()
        End Get
    End Property

    Public ReadOnly Property Stream() As String
        Get
            Return PageStream.ToString()
        End Get
    End Property
End Class
