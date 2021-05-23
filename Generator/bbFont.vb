Public Class bbFont
    Private PDFFontID As Integer
    Private PDFObject As New bbPDFObject()
    Private PDFObjectID As Integer
    Private ParentPDFObjectCollection As Collection
    Private ParentPDFFontCollection As Collection
    Private InternalFontRescale As Double
    Public Info As bbFontProgramInformation

    Public Sub New(ByRef PDFWriter As bbPDFWriter, ByVal FontID As Integer, ByVal FontName As String, ByVal FontFileName As String, ByVal FontRescale As Double)

        'Remember the parent
        ParentPDFObjectCollection = PDFWriter.GetPDFObjectsCollection()
        ParentPDFFontCollection = PDFWriter.GetPDFFontsCollection()

        'Set the scale settings
        PDFFontID = FontID
        InternalFontRescale = FontRescale

        'Read the font program file
        Info = ReadFontInformationDirectly(FontFileName)

        'Create the condensed font name
        Dim CondensedFontName As String = ""
        For i As Integer = 1 To Len(FontName)
            If Mid(FontName, i, 1) <> " " Then
                CondensedFontName &= Mid(FontName, i, 1)
            End If
        Next

        'Create the character widths string
        Dim AdvanceWidthsString As String = ""
        With Info.ANSICharacterMap
            For i As Integer = 0 To .CharacterCount - 1
                AdvanceWidthsString &= CStr(Info.HorizontalMetrics.AdvanceWidth(.CharacterToGlyphMap(i))) & " "
            Next
        End With
        AdvanceWidthsString = Trim(AdvanceWidthsString)

        'Create the BBox string
        Dim BBoxString As String = ""
        BBoxString &= Info.FontHeader.xMin
        BBoxString &= " " & Info.FontHeader.yMin
        BBoxString &= " " & Info.FontHeader.xMax
        BBoxString &= " " & Info.FontHeader.yMax

        'Write the font information to the PDF stream
        PDFObject.Write("<< /Type /Font")
        PDFObject.Write("/Subtype /TrueType")
        PDFObject.Write("/BaseFont /" & CondensedFontName)
        PDFObject.Write("/FirstChar " & 0)
        PDFObject.Write("/LastChar " & Info.ANSICharacterMap.CharacterCount - 1)
        PDFObject.Write("/Widths [" & AdvanceWidthsString & "]")
        PDFObject.Write("/FontDescriptor <<")
        PDFObject.Write("/Type /FontDescriptor")
        PDFObject.Write("/FontName /" & CondensedFontName)
        PDFObject.Write("/Flags 262178")
        PDFObject.Write("/FontBBox [" & BBoxString & "]")
        PDFObject.Write("/MissingWidth " & Info.HorizontalMetrics.AdvanceWidth(0))
        'StemV width is width of small L: 'l'
        PDFObject.Write("/StemV " & Info.HorizontalMetrics.AdvanceWidth(Info.ANSICharacterMap.CharacterToGlyphMap(Asc("l"))))
        PDFObject.Write("/CapHeight " & Info.OS2.sCapHeight)
        PDFObject.Write("/Ascent " & Info.HorizontalHeader.Ascender)
        PDFObject.Write("/Descent " & Info.HorizontalHeader.Descender)
        PDFObject.Write("/ItalicAngle " & Info.PostScript.ItalicAngle)

        'Add the reference preemptively, so we can figure out where to embed the font stream
        ParentPDFObjectCollection.Add(PDFObject)
        PDFObjectID = ParentPDFObjectCollection.Count

        PDFObject.Write("/FontFile2 " & PDFObjectStringID(PDFObjectID + 1))

        PDFObject.Write(">>")
        PDFObject.Write(">>")

        ParentPDFFontCollection.Add(Me)

        'Write embedded font object stream
        Dim EmbeddedFontObj As New bbPDFObject()
        ParentPDFObjectCollection.Add(EmbeddedFontObj)
        Dim EmbeddedFontObjID As Integer = ParentPDFObjectCollection.Count
        EmbeddedFontObj.Write("<</Filter /ASCIIHexDecode")
        EmbeddedFontObj.Write("/Length " & Info.EmbeddedASCIIHexFontProgram.FontProgramHexChars)
        EmbeddedFontObj.Write("/Length1 " & Info.EmbeddedASCIIHexFontProgram.FontProgramBytes)
        EmbeddedFontObj.Write(">>")
        EmbeddedFontObj.Write("stream")
        EmbeddedFontObj.Write(Info.EmbeddedASCIIHexFontProgram.FontProgramASCIIHex)
        EmbeddedFontObj.Write("endstream")
    End Sub

    Public Function Scale() As Double
        Return InternalFontRescale
    End Function

    Public Function GetPDFID() As String
        Return "/F" & PDFFontID
    End Function

    Public Function GetPDFObjectID() As String
        Return CStr(PDFObjectID & " 0 R")
    End Function
End Class
