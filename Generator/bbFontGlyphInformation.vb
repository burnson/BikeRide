Imports System.Runtime.InteropServices
Imports System.Text
Imports System.Drawing

Public Module bbFontGlyphInformation
    Private Declare Function GetCharABCWidths Lib "gdi32" Alias "GetCharABCWidthsA" (ByVal hdc As Int32, ByVal uFirstChar As Int32, ByVal uLastChar As Int32, ByRef lpabc As ABC) As Int32
    Private Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Int32, ByVal uFirstChar As Int32, ByVal uLastChar As Int32, ByRef lpabc As Int32) As Int32
    Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal H As Int32, ByVal W As Int32, ByVal E As Int32, ByVal O As Int32, ByVal W As Int32, ByVal i As Int32, ByVal u As Int32, ByVal S As Int32, ByVal C As Int32, ByVal OP As Int32, ByVal CP As Int32, ByVal Q As Int32, ByVal PAF As Int32, ByVal F As String) As Int32
    Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" (ByVal hdc As Int32, ByRef lpMetrics As TEXTMETRIC) As Int32
    Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Int32, ByVal hdc As Int32) As Int32
    Private Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Int32, ByRef lpStr As Byte, ByVal nCount As Int32, ByRef lpRect As RECT, ByVal wFormat As Int32) As Int32
    Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Int32, ByVal hObject As Int32) As Int32
    Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Int32, ByRef lpRect As RECT) As Int32
    Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Int32, ByVal index As Int32) As Int32
    Private Declare Function GetOutlineTextMetrics Lib "gdi32" Alias "GetOutlineTextMetricsA" (ByVal hdc As Int32, ByVal Length As Int32, ByRef Data As OUTLINETEXTMETRIC) As Int32
    Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Int32) As Int32
    Private Declare Function CreateDC Lib "gdi32" Alias "CreateDCA" (ByRef DriverName As Byte, ByRef DeviceName As Byte, ByVal Null As IntPtr, ByVal Null2 As IntPtr) As Int32
    Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Int32) As Int32

    Private Const TMPF_TRUETYPE = 4
    Private Const FW_DONTCARE = 0
    Private Const DEFAULT_CHARSET = 1
    Private Const OUT_DEFAULT_PRECIS = 0
    Private Const CLIP_DEFAULT_PRECIS = 0
    Private Const DEFAULT_QUALITY = 0
    Private Const FF_DONTCARE = 0    '  Don't care or don't know.
    Private Const DT_LEFT = &H0
    Private Const LF_FACESIZE = 32
    Private Const LOGPIXELSY = 90
    Private Const FW_SEMIBOLD = 600
    Private Const FW_NORMAL = 400
    Private Const FW_BOLD = 700


    Private Structure ABC
        Public abcA As Int32
        Public abcB As Int32
        Public abcC As Int32
    End Structure

    Private Structure RECT
        Public Left As Int32
        Public Top As Int32
        Public Right As Int32
        Public Bottom As Int32
    End Structure

    Private Structure TEXTMETRIC
        Public tmHeight As Int32
        Public tmAscent As Int32
        Public tmDescent As Int32
        Public tmInternalLeading As Int32
        Public tmExternalLeading As Int32
        Public tmAveCharWidth As Int32
        Public tmMaxCharWidth As Int32
        Public tmWeight As Int32
        Public tmOverhang As Int32
        Public tmDigitizedAspectX As Int32
        Public tmDigitizedAspectY As Int32
        Public tmFirstChar As Byte
        Public tmLastChar As Byte
        Public tmDefaultChar As Byte
        Public tmBreakChar As Byte
        Public tmItalic As Byte
        Public tmUnderlined As Byte
        Public tmStruckOut As Byte
        Public tmPitchAndFamily As Byte
        Public tmCharSet As Byte
    End Structure



    Private Structure PANOSE
        Public bFamilyType As Byte
        Public bSerifStyle As Byte
        Public bWeight As Byte
        Public bProportion As Byte
        Public bContrast As Byte
        Public bStrokeVariation As Byte
        Public bArmStyle As Byte
        Public bLetterform As Byte
        Public bMidline As Byte
        Public bXHeight As Byte
    End Structure

    Private Structure OUTLINETEXTMETRIC
        Public otmSize As Int32
        Public otmTextMetrics As TEXTMETRIC
        Public otmFiller As Byte
        Public otmPanoseNumber As PANOSE
        Public otmfsSelection As Int32
        Public otmfsType As Int32
        Public otmsCharSlopeRise As Int32
        Public otmsCharSlopeRun As Int32
        Public otmItalicAngle As Int32
        Public otmEMSquare As Int32
        Public otmAscent As Int32
        Public otmDescent As Int32
        Public otmLineGap As Int32
        Public otmsCapEmHeight As Int32
        Public otmsXHeight As Int32
        Public otmrcFontBox As RECT
        Public otmMacAscent As Int32
        Public otmMacDescent As Int32
        Public otmMacLineGap As Int32
        Public otmusMinimumPPEM As Int32
        Public otmptSubscriptSize As Win32POINT
        Public otmptSubscriptOffset As Win32POINT
        Public otmptSuperscriptSize As Win32POINT
        Public otmptSuperscriptOffset As Win32POINT
        Public otmsStrikeoutSize As Int32
        Public otmsStrikeoutPosition As Int32
        Public otmsUnderscoreSize As Int32
        Public otmsUnderscorePosition As Int32
        Public otmpFamilyName As System.IntPtr
        Public otmpFaceName As System.IntPtr
        Public otmpStyleName As System.IntPtr
        Public otmpFullName As System.IntPtr
    End Structure

    Private Structure Win32POINT
        Public x As Int32
        Public y As Int32
    End Structure

    Public Structure bbGlyphInformation
        Public CondensedFontName As String
        Public Ascent As Long
        Public Descent As Long
        Public CapHeight As Long
        Public FirstCharacter As Long
        Public LastCharacter As Long
        Public CharacterWidths() As Long
        Public CharacterWidthsString As String
        Public GlobalBoundingBox As Rectangle
        Public GlobalBoundingBoxString As String
    End Structure

    Private Function GetByteString(ByRef Str As String) As Byte()
        Dim ByteData(Len(Str)) As Byte
        For i As Long = 0 To Len(Str)
            If i = Len(Str) Then
                ByteData(i) = 0
            Else
                ByteData(i) = Asc(Mid(Str, i + 1, 1))
            End If
        Next
        Return ByteData
    End Function

    Public Function GetCharacterWidths(ByVal FontName As String, ByVal FontBold As Boolean, ByVal FontItalic As Boolean) As bbGlyphInformation
        Const arbitraryPointSize As Int32 = 50
        Dim GlyphInformation As bbGlyphInformation
        GlyphInformation.CondensedFontName = ""
        For i As Long = 1 To Len(FontName)
            If Mid(FontName, i, 1) <> " " Then
                GlyphInformation.CondensedFontName &= Mid(FontName, i, 1)
            End If
        Next

        Dim Null As IntPtr = 0
        Dim hDC As Int32 = CreateDC(GetByteString("display")(0), GetByteString("")(0), Null, Null)
        Dim hMemDC = CreateCompatibleDC(hDC)
        Dim Res As Int32 = GetDeviceCaps(hMemDC, LOGPIXELSY)
        Dim nHeight As Int32 = -arbitraryPointSize * Res / 72
        Dim BoldValue As Int32 = FW_NORMAL
        If FontBold Then BoldValue = FW_SEMIBOLD
        Dim ItalicsValue As Int32 = 0
        If FontItalic Then ItalicsValue = -1
        Dim hFont As Int32 = CreateFont(nHeight, 0, 0, 0, BoldValue, ItalicsValue, 0, 0, DEFAULT_CHARSET, OUT_DEFAULT_PRECIS, CLIP_DEFAULT_PRECIS, DEFAULT_QUALITY, FF_DONTCARE, FontName)
        Dim hOldFont As Int32 = SelectObject(hMemDC, hFont)

        Dim tMetrics As TEXTMETRIC
        Dim otMetrics As OUTLINETEXTMETRIC

        GetTextMetrics(hMemDC, tMetrics)

        Dim FirstCharacter As Int32 = tMetrics.tmFirstChar
        Dim LastCharacter As Int32 = tMetrics.tmLastChar
        GlyphInformation.FirstCharacter = FirstCharacter
        GlyphInformation.LastCharacter = LastCharacter

        Dim bTrueType As Boolean = (tMetrics.tmPitchAndFamily = TMPF_TRUETYPE)

        If bTrueType Then
            otMetrics.otmSize = Marshal.SizeOf(otMetrics)
            GetOutlineTextMetrics(hMemDC, otMetrics.otmSize, otMetrics)
        End If

        ReDim GlyphInformation.CharacterWidths(LastCharacter)
        GlyphInformation.CharacterWidthsString = ""
        For c As Integer = FirstCharacter To LastCharacter
            Dim Ret As Int32
            Dim ABC As ABC
            Dim Width As Int32
            If bTrueType Then
                Ret = GetCharABCWidths(hMemDC, c, c, ABC)
                Width = ABC.abcA + ABC.abcB + ABC.abcC
            Else
                Ret = GetCharWidth32(hMemDC, c, c, Width)
                Width = (1000 / arbitraryPointSize) * (72 * Width / Res) 'rescale to 1000 points
            End If
            GlyphInformation.CharacterWidths(c) = Width
            GlyphInformation.CharacterWidthsString &= CStr(Width) & " "
        Next

        GlyphInformation.CharacterWidthsString = Trim(GlyphInformation.CharacterWidthsString)


        GlyphInformation.GlobalBoundingBox.X = otMetrics.otmrcFontBox.Left
        GlyphInformation.GlobalBoundingBox.Y = otMetrics.otmrcFontBox.Top
        GlyphInformation.GlobalBoundingBox.Width = otMetrics.otmrcFontBox.Right - otMetrics.otmrcFontBox.Left
        GlyphInformation.GlobalBoundingBox.Height = otMetrics.otmrcFontBox.Bottom - otMetrics.otmrcFontBox.Top

        With GlyphInformation
            .GlobalBoundingBoxString = .GlobalBoundingBox.Left & " " & .GlobalBoundingBox.Top & " " & .GlobalBoundingBox.Right & " " & .GlobalBoundingBox.Bottom
        End With

        GlyphInformation.CapHeight = otMetrics.otmsCapEmHeight
        GlyphInformation.Ascent = otMetrics.otmAscent
        GlyphInformation.Descent = otMetrics.otmDescent

        SelectObject(hMemDC, hOldFont)
        DeleteDC(hMemDC)
        DeleteDC(hDC)
        Return GlyphInformation
    End Function

    '*****************READ DIRECTLY FROM THE FONT FILE*********************
    Public Const Shift0 As UInt64 = 1
    Public Const Shift1 As UInt64 = 256
    Public Const Shift2 As UInt64 = 65536
    Public Const Shift3 As UInt64 = Shift1 * Shift1 * Shift1

    'BIG-ENDIAN READING
    Private Function ReadBYTE(ByVal Bytes() As Byte, ByRef Position As Long) As Byte
        Dim Sum As UInt64 = Bytes(Position + 0) * Shift0
        Position += 1
        Return CByte(Sum)
    End Function

    Private Function ReadUSHORT(ByVal Bytes() As Byte, ByRef Position As Long) As UInt16
        Dim Sum As UInt64 = Bytes(Position + 0) * Shift1 + Bytes(Position + 1) * Shift0
        Position += 2
        Return Sum
    End Function

    Private Function ReadULONG(ByVal Bytes() As Byte, ByRef Position As Long) As UInt32
        Dim Sum As UInt64 = Bytes(Position + 0) * Shift3 + Bytes(Position + 1) * Shift2 + Bytes(Position + 2) * Shift1 + Bytes(Position + 3) * Shift0
        Position += 4
        Return Sum
    End Function

    Private Function ReadTAG(ByVal Bytes() As Byte, ByRef Position As Long) As String
        Dim Concatenation As String = Chr(Bytes(Position + 0)) & Chr(Bytes(Position + 1)) & Chr(Bytes(Position + 2)) & Chr(Bytes(Position + 3))
        Position += 4
        Return Concatenation
    End Function

    Private Function ReadSHORT(ByVal Bytes() As Byte, ByRef Position As Long) As Int16
        Dim x As Int64 = ReadUSHORT(Bytes, Position)
        If x >= 32768 Then
            x -= 65536
        End If
        Return x
    End Function

    Private Function ReadLONG(ByVal Bytes() As Byte, ByRef Position As Long) As Int32
        Dim x As Int64 = ReadULONG(Bytes, Position)
        If x >= CLng(2 ^ 31) Then
            x -= CLng(2 ^ 32)
        End If
        Return x
    End Function

    Private Function ReadCHAR(ByVal Bytes() As Byte, ByRef Position As Long) As Int16
        Dim x As Int16 = ReadBYTE(Bytes, Position)
        If x >= 128 Then
            x -= 256
        End If
        Return x
    End Function

    Private Function ReadFWORD(ByVal Bytes() As Byte, ByRef Position As Long) As Int16
        Return ReadSHORT(Bytes, Position)
    End Function

    Private Function ReadUFWORD(ByVal Bytes() As Byte, ByRef Position As Long) As Int16
        Return ReadUSHORT(Bytes, Position)
    End Function

    Private Function ReadFIXED(ByVal Bytes() As Byte, ByRef Position As Long) As Double
        Dim IntegerPart As Double = ReadSHORT(Bytes, Position)
        Dim FractionalPart As Double = ReadUSHORT(Bytes, Position)
        Dim FixedPointNumber As Double = IntegerPart + Math.Sign(IntegerPart) * (FractionalPart / 65536.0)
        Return FixedPointNumber
    End Function

    Public Structure TableDirectoryEntry
        Public tag As String
        Public checkSum As Long
        Public offset As Long
        Public length As Long
    End Structure

    Public Structure bbTableFontHeader
        Public versionNumber As Long
        Public fontRevision As Long
        Public checkSumAdjustment As Long
        Public magicNumber As Long
        Public flags As Long
        Public unitsPerEm As UInt16
        Public date_created As UInt64
        Public date_modified As UInt64
        Public xMin As Long
        Public yMin As Long
        Public xMax As Long
        Public yMax As Long
        Public macStyle As Long
        Public lowestRecPPEM As Long
        Public fontDirectionHint As Long
        Public indexToLocFormat As Long
        Public glyphDataFormat As Long
    End Structure

    Public Structure bbTableHorizontalHeader
        Public versionNumber As Long
        Public Ascender As Long
        Public Descender As Long
        Public LineGap As Long
        Public advanceWidthMax As Long
        Public minLeftSideBearing As Long
        Public minRightSideBearing As Long
        Public xMaxExtent As Long
        Public caretSlopeRise As Long
        Public caretSlopeRun As Long
        Public caretOffset As Long
        Public Reserved1 As Long
        Public Reserved2 As Long
        Public Reserved3 As Long
        Public Reserved4 As Long
        Public metricDataFormat As Long
        Public numberOfHMetrics As Long
    End Structure

    Public Structure bbTableHorizontalMetrics
        Public AdvanceWidth() As Long
        Public LeftSideBearing() As Long
        Public numberofHMetrics As Long
    End Structure

    Public Structure bbTableANSICharacterMap
        Public CharacterToGlyphMap() As Long
        Public CharacterCount As Long
    End Structure

    Public Structure bbTableKerningPairs
        Public KerningPairCount As Long
        Public GlyphLeft() As Long
        Public GlyphRight() As Long
        Public KernValue() As Long
        Public Horizontal As Boolean
    End Structure

    Public Structure bbTablePostScript
        Public ItalicAngle As Double
        Public underlinePosition As Long
        Public underlineThickness As Long
        Public isFixedPitch As Boolean
    End Structure

    Public Structure bbTableOS2
        Public xAvgCharWidth As Long
        Public usWeightClass As Long
        Public usWidthClass As Long
        Public fsType As Long
        Public ySubscriptXSize As Long
        Public ySubscriptYSize As Long
        Public ySubscriptXOffset As Long
        Public ySubscriptYOffset As Long
        Public ySuperscriptXSize As Long
        Public ySuperscriptYSize As Long
        Public ySuperscriptXOffset As Long
        Public ySuperscriptYOffset As Long
        Public yStrikeoutSize As Long
        Public yStrikeoutPosition As Long
        Public sFamilyClass As Long
        Public ulUnicodeRange1 As ULong
        Public ulUnicodeRange2 As ULong
        Public ulUnicodeRange3 As ULong
        Public ulUnicodeRange4 As ULong
        Public fsSelection As Long
        Public usFirstCharIndex As Long
        Public usLastCharIndex As Long
        Public sTypoAscender As Long
        Public sTypoDescender As Long
        Public sTypoLineGap As Long
        Public usWinAscent As Long
        Public usWinDescent As Long
        Public ulCodePageRange1 As ULong
        Public ulCodePageRange2 As ULong
        Public sxHeight As Long
        Public sCapHeight As Long
        Public usDefaultChar As Long
        Public usBreakChar As Long
        Public usMaxContext As Long
    End Structure

    Public Structure bbASCIIHexFontProgram
        Public FontProgramBytes As Long
        Public FontProgramHexChars As Long
        Public FontProgramASCIIHex As String
    End Structure

    Public Structure bbFontProgramInformation
        Public FontHeader As bbTableFontHeader
        Public HorizontalHeader As bbTableHorizontalHeader
        Public HorizontalMetrics As bbTableHorizontalMetrics
        Public ANSICharacterMap As bbTableANSICharacterMap
        Public Kerning As bbTableKerningPairs
        Public PostScript As bbTablePostScript
        Public OS2 As bbTableOS2
        Public EmbeddedASCIIHexFontProgram As bbASCIIHexFontProgram
    End Structure

    Private Function CreateEmbeddedASCIIHexFontProgram(ByRef Bytes() As Byte, ByVal ByteCount As Long) As bbASCIIHexFontProgram
        Dim x As bbASCIIHexFontProgram
        With x
            .FontProgramBytes = ByteCount
            .FontProgramHexChars = ByteCount * 2 + 1
            .FontProgramASCIIHex = ""
            Dim out() As Byte
            ReDim out(.FontProgramHexChars - 1)

            Dim a As Integer, b As Integer, c As Integer, j As Integer, k As Integer

            For i As Integer = 0 To ByteCount - 1
                c = Bytes(i)
                a = c \ 16 + 48
                b = c Mod 16 + 48
                j = i * 2
                k = j + 1
                If a >= 58 Then a += 7
                If b >= 58 Then b += 7
                out(j) = a
                out(k) = b
            Next
            out(k + 1) = Asc(">")
            .FontProgramASCIIHex = System.Text.Encoding.ASCII.GetString(out)
        End With
        Return x
    End Function

    Private Function ReadOS2Table(ByRef Bytes() As Byte, ByRef Position As Long) As bbTableOS2
        Dim x As bbTableOS2
        With x
            ReadUSHORT(Bytes, Position) 'version
            .xAvgCharWidth = ReadSHORT(Bytes, Position)
            .usWeightClass = ReadUSHORT(Bytes, Position)
            .usWidthClass = ReadUSHORT(Bytes, Position)
            .fsType = ReadUSHORT(Bytes, Position)
            .ySubscriptXSize = ReadSHORT(Bytes, Position)
            .ySubscriptYSize = ReadSHORT(Bytes, Position)
            .ySubscriptXOffset = ReadSHORT(Bytes, Position)
            .ySubscriptYOffset = ReadSHORT(Bytes, Position)
            .ySuperscriptXSize = ReadSHORT(Bytes, Position)
            .ySuperscriptYSize = ReadSHORT(Bytes, Position)
            .ySuperscriptXOffset = ReadSHORT(Bytes, Position)
            .ySuperscriptYOffset = ReadSHORT(Bytes, Position)
            .yStrikeoutSize = ReadSHORT(Bytes, Position)
            .yStrikeoutPosition = ReadSHORT(Bytes, Position)
            .sFamilyClass = ReadSHORT(Bytes, Position)
            Position += 10 'PANOSE
            .ulUnicodeRange1 = ReadULONG(Bytes, Position)
            .ulUnicodeRange2 = ReadULONG(Bytes, Position)
            .ulUnicodeRange3 = ReadULONG(Bytes, Position)
            .ulUnicodeRange4 = ReadULONG(Bytes, Position)
            Position += 4 'VENDOR ID
            .fsSelection = ReadUSHORT(Bytes, Position)
            .usFirstCharIndex = ReadUSHORT(Bytes, Position)
            .usLastCharIndex = ReadUSHORT(Bytes, Position)
            .sTypoAscender = ReadSHORT(Bytes, Position)
            .sTypoDescender = ReadSHORT(Bytes, Position)
            .sTypoLineGap = ReadSHORT(Bytes, Position)
            .usWinAscent = ReadUSHORT(Bytes, Position)
            .usWinDescent = ReadUSHORT(Bytes, Position)
            .ulCodePageRange1 = ReadULONG(Bytes, Position)
            .ulCodePageRange2 = ReadULONG(Bytes, Position)
            .sxHeight = ReadSHORT(Bytes, Position)
            .sCapHeight = ReadSHORT(Bytes, Position)
            .usDefaultChar = ReadSHORT(Bytes, Position)
            .usBreakChar = ReadSHORT(Bytes, Position)
            .usMaxContext = ReadSHORT(Bytes, Position)
        End With
        Return x
    End Function

    Private Function ReadPostScriptTable(ByRef Bytes() As Byte, ByRef Position As Long) As bbTablePostScript
        Dim x As bbTablePostScript
        With x
            Dim Version As Long = ReadULONG(Bytes, Position)
            .ItalicAngle = ReadFIXED(Bytes, Position)
            .underlinePosition = ReadFWORD(Bytes, Position)
            .underlineThickness = ReadFWORD(Bytes, Position)
            .isFixedPitch = (ReadULONG(Bytes, Position) <> 0)
        End With
        Return x
    End Function

    Private Function ReadKerningPairsTable(ByRef Bytes() As Byte, ByRef Position As Long) As bbTableKerningPairs
        Dim x As bbTableKerningPairs
        With x
            .KerningPairCount = 0
            ReDim .GlyphLeft(0)
            ReDim .GlyphRight(0)
            ReDim .KernValue(0)
        End With

        Dim TableVersion As Long = ReadUSHORT(Bytes, Position)
        Dim NumTables As Long = ReadUSHORT(Bytes, Position)

        For i As Long = 0 To NumTables - 1
            Dim SubtableVersion As Long = ReadUSHORT(Bytes, Position)
            Dim Length As Long = ReadUSHORT(Bytes, Position)
            Dim Coverage As UInt16 = ReadUSHORT(Bytes, Position)
            If Coverage And 1 Then
                'horizontal
                x.Horizontal = True
            Else
                x.Horizontal = False
            End If

            If Not (Coverage And 2) And (Coverage \ 256) = 0 Then
                x.KerningPairCount = ReadUSHORT(Bytes, Position)
                ReDim x.GlyphLeft(x.KerningPairCount - 1)
                ReDim x.GlyphRight(x.KerningPairCount - 1)
                ReDim x.KernValue(x.KerningPairCount - 1)

                Dim searchRange As Long = ReadUSHORT(Bytes, Position)
                Dim entrySelector As Long = ReadUSHORT(Bytes, Position)
                Dim rangeShift As Long = ReadUSHORT(Bytes, Position)

                For j As Long = 0 To x.KerningPairCount - 1
                    With x
                        .GlyphLeft(j) = ReadUSHORT(Bytes, Position)
                        .GlyphRight(j) = ReadUSHORT(Bytes, Position)
                        .KernValue(j) = ReadFWORD(Bytes, Position)
                    End With
                Next
            End If
        Next

        Return x
    End Function

    Private Function ReadANSICharacterMapTable(ByRef Bytes() As Byte, ByRef Position As Long) As bbTableANSICharacterMap
        Dim x As bbTableANSICharacterMap
        x.CharacterCount = 0
        ReDim x.CharacterToGlyphMap(0)

        Dim OriginalPosition As Long = Position
        Dim cmapVersionNumber As Long = ReadUSHORT(Bytes, Position)
        Dim numTables As Long = ReadUSHORT(Bytes, Position)

        Dim platformID As Long
        Dim encodingID As Long
        Dim offset As Long
        For i As Long = 0 To numTables - 1
            platformID = ReadUSHORT(Bytes, Position)
            encodingID = ReadUSHORT(Bytes, Position)
            offset = ReadULONG(Bytes, Position)

            If platformID = 3 And encodingID = 1 Then
                'Found the Microsoft UNICODE encoding table
                Exit For
            End If
        Next

        'Go to the subtable
        Position = OriginalPosition + offset

        Dim Format As Long = ReadUSHORT(Bytes, Position)
        If Format <> 4 Then Return x 'Make sure it's the Microsoft segment to delta mapping

        Dim Length As Long = ReadUSHORT(Bytes, Position)
        Dim Language As Long = ReadUSHORT(Bytes, Position)
        Dim segCountX2 As Long = ReadUSHORT(Bytes, Position)
        Dim searchRange As Long = ReadUSHORT(Bytes, Position)
        Dim entrySelector As Long = ReadUSHORT(Bytes, Position)
        Dim rangeShift As Long = ReadUSHORT(Bytes, Position)
        Dim segCount As Long = segCountX2 / 2
        Dim endCharacterCode(segCount - 1) As Long
        For i As Long = 0 To segCount - 1
            endCharacterCode(i) = ReadUSHORT(Bytes, Position)
        Next
        Dim reservedPad As Long = ReadUSHORT(Bytes, Position)
        Dim startCharacterCode(segCount - 1) As Long
        For i As Long = 0 To segCount - 1
            startCharacterCode(i) = ReadUSHORT(Bytes, Position)
        Next
        Dim idDelta(segCount - 1) As Long
        For i As Long = 0 To segCount - 1
            idDelta(i) = ReadSHORT(Bytes, Position)
        Next
        Dim idRangeOffset(segCount - 1) As Long
        For i As Long = 0 To segCount - 1
            idRangeOffset(i) = ReadUSHORT(Bytes, Position)
        Next

        'Resize the structure's array
        x.CharacterCount = 256
        ReDim x.CharacterToGlyphMap(x.CharacterCount - 1)

        'Perform a search for each of the characters, and do a mapping for each
        Dim glyphIDArrayPosition As Long = Position
        Dim glyphIndex As Long = 0
        Dim characterIndex As Long = 0
        Dim segmentIndex As Long = 0
        For characterIndex = 0 To x.CharacterCount - 1
            'Search for the character index
            segmentIndex = -1
            For j As Long = 0 To segCount - 1
                If endCharacterCode(j) <> 65535 Then
                    If characterIndex <= endCharacterCode(j) Then
                        If startCharacterCode(j) <= characterIndex Then
                            segmentIndex = j
                            Exit For
                        End If
                    End If
                End If
            Next

            'Found character in mapping, now determine it's glyph index
            If segmentIndex >= 0 Then
                If idRangeOffset(segmentIndex) = 0 Then
                    glyphIndex = idDelta(segmentIndex) + characterIndex
                Else
                    Dim CharacterCodeOffset As Long = (characterIndex - startCharacterCode(segmentIndex))
                    Dim NumberOfUShortsOffset As Long = idRangeOffset(segmentIndex) / 2
                    Dim NumberOfUShortsInRangeRemaining As Long = -(segCount - segmentIndex)
                    Dim glyphIDAddress = glyphIDArrayPosition + (NumberOfUShortsInRangeRemaining + NumberOfUShortsOffset + CharacterCodeOffset) * 2

                    glyphIndex = ReadUSHORT(Bytes, glyphIDAddress)
                End If
            Else
                glyphIndex = 0
            End If
            x.CharacterToGlyphMap(characterIndex) = glyphIndex
        Next
        Return x
    End Function

    Private Function ReadHorizontalMatrixTable(ByVal hhea As bbTableHorizontalHeader, ByRef Bytes() As Byte, ByRef Position As Long) As bbTableHorizontalMetrics
        Dim x As bbTableHorizontalMetrics
        With x
            .numberofHMetrics = hhea.numberOfHMetrics
            ReDim .AdvanceWidth(.numberofHMetrics * 2 - 1)
            ReDim .LeftSideBearing(.numberofHMetrics * 2 - 1)
            For i As Long = 0 To .numberofHMetrics - 1
                .AdvanceWidth(i) = ReadUSHORT(Bytes, Position) / 2 '888888
                .LeftSideBearing(i) = ReadSHORT(Bytes, Position)
            Next
        End With
        Return x
    End Function

    Private Function ReadHorizontalHeaderTable(ByRef Bytes() As Byte, ByRef Position As Long) As bbTableHorizontalHeader
        Dim x As bbTableHorizontalHeader
        With x
            .versionNumber = ReadULONG(Bytes, Position)
            .Ascender = ReadSHORT(Bytes, Position)
            .Descender = ReadSHORT(Bytes, Position)
            .LineGap = ReadSHORT(Bytes, Position)
            .advanceWidthMax = ReadUSHORT(Bytes, Position)
            .minLeftSideBearing = ReadSHORT(Bytes, Position)
            .minRightSideBearing = ReadSHORT(Bytes, Position)
            .xMaxExtent = ReadSHORT(Bytes, Position)
            .caretSlopeRise = ReadSHORT(Bytes, Position)
            .caretSlopeRun = ReadSHORT(Bytes, Position)
            .caretOffset = ReadSHORT(Bytes, Position)
            .Reserved1 = ReadSHORT(Bytes, Position)
            .Reserved2 = ReadSHORT(Bytes, Position)
            .Reserved3 = ReadSHORT(Bytes, Position)
            .Reserved4 = ReadSHORT(Bytes, Position)
            .metricDataFormat = ReadSHORT(Bytes, Position)
            .numberOfHMetrics = ReadUSHORT(Bytes, Position)
        End With
        Return x
    End Function

    Private Function ReadFontHeaderTable(ByVal Bytes() As Byte, ByRef Position As Long) As bbTableFontHeader
        Dim x As bbTableFontHeader
        With x
            .versionNumber = ReadULONG(Bytes, Position)
            .fontRevision = ReadULONG(Bytes, Position)
            .checkSumAdjustment = ReadULONG(Bytes, Position)
            .magicNumber = ReadULONG(Bytes, Position)
            .flags = ReadUSHORT(Bytes, Position)
            .unitsPerEm = ReadUSHORT(Bytes, Position)
            ReadLONG(Bytes, Position)
            ReadLONG(Bytes, Position)
            ReadLONG(Bytes, Position)
            ReadLONG(Bytes, Position)
            .xMin = ReadSHORT(Bytes, Position)
            .yMin = ReadSHORT(Bytes, Position)
            .xMax = ReadSHORT(Bytes, Position)
            .yMax = ReadSHORT(Bytes, Position)
            .macStyle = ReadUSHORT(Bytes, Position)
            .lowestRecPPEM = ReadUSHORT(Bytes, Position)
            .fontDirectionHint = ReadSHORT(Bytes, Position)
            .indexToLocFormat = ReadSHORT(Bytes, Position)
            .glyphDataFormat = ReadSHORT(Bytes, Position)
        End With
        Return x
    End Function

    Private Function ReadTableDirectoryEntry(ByVal Bytes() As Byte, ByRef Position As Long) As TableDirectoryEntry
        Dim Entry As TableDirectoryEntry
        With Entry
            .tag = ReadTAG(Bytes, Position)
            .checkSum = ReadULONG(Bytes, Position)
            .offset = ReadULONG(Bytes, Position)
            .length = ReadULONG(Bytes, Position)
        End With
        Return Entry
    End Function

    Public Function ReadFontInformationDirectly(ByVal FontFile As String) As bbFontProgramInformation
        'Debug.Print("ReadFontInformationDirectly(""" & FontFile & """)")
        'Debug.Print("Loading file into memory...")
        Dim Bytes() As Byte
        Bytes = My.Computer.FileSystem.ReadAllBytes(FontFile)
        Dim ByteCount As Long = UBound(Bytes) - LBound(Bytes) + 1
        Dim Position As Long = 0
        'Debug.Print("File contains TTF header: " & (ReadULONG(Bytes, 0) = 65536))

        'Debug.Print("Reading data...")

        '=====OFFSET TABLE=====
        'Debug.Print("OFFSET TABLE")

        ReadULONG(Bytes, Position) 'sfnt version:  0x00010000 for version 1.0 (65536)
        Dim numTables As Long = ReadUSHORT(Bytes, Position) 'numTables: Number of tables
        Dim searchRange As Long = ReadUSHORT(Bytes, Position) 'searchRange: (Maximum power of 2 <= numTables) x 16
        Dim entrySelector As Long = ReadUSHORT(Bytes, Position) 'entrySelector: Log2(maximum power of 2 <= numTables)
        Dim rangeShift As Long = ReadUSHORT(Bytes, Position) 'rangeShift: NumTables x 16-searchRange

        'Debug.Indent()
        'Debug.Print("numTables: " & numTables)
        'Debug.Print("searchRange: " & searchRange)
        'Debug.Print("entrySelector: " & entrySelector)
        'Debug.Print("rangeShift: " & rangeShift)
        'Debug.Unindent()

        '======TABLE DIRECTORY=====
        Dim headStart As Long
        Dim hheaStart As Long
        Dim hmtxStart As Long
        Dim cmapStart As Long
        Dim kernStart As Long
        Dim postStart As Long
        Dim OS2Start As Long

        'Debug.Print("TABLE DIRECTORY")
        'Debug.Indent()
        Dim Tables(0) As TableDirectoryEntry
        If numTables > 0 Then
            ReDim Tables(numTables - 1)
            For i As Long = 0 To numTables - 1
                Tables(i) = ReadTableDirectoryEntry(Bytes, Position)
                If Tables(i).tag = "head" Then headStart = Tables(i).offset
                If Tables(i).tag = "hhea" Then hheaStart = Tables(i).offset
                If Tables(i).tag = "hmtx" Then hmtxStart = Tables(i).offset
                If Tables(i).tag = "cmap" Then cmapStart = Tables(i).offset
                If Tables(i).tag = "kern" Then kernStart = Tables(i).offset
                If Tables(i).tag = "post" Then postStart = Tables(i).offset
                If Tables(i).tag = "OS/2" Then OS2Start = Tables(i).offset

                'Debug.Print("Tag #" & i & ": " & Tables(i).tag)
            Next
        End If
        'Debug.Unindent()

        'Important tables:
        'head -- font header
        'hhea -- horizontal header
        'hmtx -- horizontal metrics (advance widths/left side bearings)
        'OS/2 -- OS specific metrics
        'post -- postscript information

        '=====FONT HEADER TABLE (head)=====
        'Debug.Print("TABLE (head)")
        Position = headStart
        Dim head As bbTableFontHeader = ReadFontHeaderTable(Bytes, Position)
        'Debug.Indent()
        'Debug.Print("units per Em: " & head.unitsPerEm)
        'Debug.Print("xMin: " & head.xMin)
        'Debug.Print("yMin: " & head.yMin)
        'Debug.Print("xMax: " & head.xMax)
        'Debug.Print("yMax: " & head.yMax)
        'Debug.Unindent()

        '=====HORIZONTAL HEADER TABLE (hhea)=====
        'Debug.Print("TABLE (hhea)")
        Position = hheaStart
        Dim hhea As bbTableHorizontalHeader = ReadHorizontalHeaderTable(Bytes, Position)
        'Debug.Indent()
        'Debug.Print("numberOfHMetrics: " & hhea.numberOfHMetrics)
        'Debug.Print("minLeftSideBearing: " & hhea.minLeftSideBearing)
        'Debug.Print("minRightSideBearing: " & hhea.minRightSideBearing)
        'Debug.Print("xMaxExtent: " & hhea.xMaxExtent)
        'Debug.Print("Ascender: " & hhea.Ascender)
        'Debug.Print("Descender: " & hhea.Descender)
        'Debug.Print("advanceWidthMax: " & hhea.advanceWidthMax)
        'Debug.Unindent()

        '=====HORIZONTAL MATRIX TABLE (hmtx)=====
        'Debug.Print("TABLE (hmtx)")
        Position = hmtxStart
        Dim hmtx As bbTableHorizontalMetrics = ReadHorizontalMatrixTable(hhea, Bytes, Position)
        'Debug.Indent()
        For i As Long = 0 To hmtx.numberofHMetrics - 1
            'hmtx.AdvanceWidth(i) = Fix(CDbl(hmtx.AdvanceWidth(i)) * CDbl(head.unitsPerEm) / 1000.0)
            'Debug.Print("advanceWidth[" & i & "] = " & hmtx.AdvanceWidth(i))
        Next i
        'Debug.Unindent()

        '=====CHARACTER MAPPING TABLE (cmap)=====
        'Debug.Print("TABLE (cmap)")
        Position = cmapStart
        'Debug.Indent()
        'Debug.Print("Reading character mappings...")
        Dim cmap As bbTableANSICharacterMap = ReadANSICharacterMapTable(Bytes, Position)
        'Debug.Print("Advance widths mapped as characters codes:")
        For i As Long = 0 To cmap.CharacterCount - 1
            Dim advanceWidth As Long = hmtx.AdvanceWidth(cmap.CharacterToGlyphMap(i))
            'Debug.Print("advanceWidth[" & i & "] = " & advanceWidth)
        Next
        'Debug.Unindent()

        '=====KERNING PAIRS TABLE (kern)=====
        'Debug.Print("TABLE (kern)")
        Position = kernStart
        'Debug.Indent()
        Dim kern As bbTableKerningPairs = ReadKerningPairsTable(Bytes, Position)
        For i As Long = 0 To kern.KerningPairCount - 1
            'Debug.Print("Pair #" & i & ": L(" & kern.GlyphLeft(i) & ")-R(" & kern.GlyphRight(i) & "): " & kern.KernValue(i))
        Next
        'Debug.Unindent()

        '=====POSTSCRIPT TABLE (post)=====
        Position = postStart
        Dim post As bbTablePostScript = ReadPostScriptTable(Bytes, Position)

        '=====OS/2 TABLE (os/2)======
        Position = OS2Start
        Dim os2 As bbTableOS2 = ReadOS2Table(Bytes, Position)

        'Create a font program information structure...
        Dim x As bbFontProgramInformation
        With x
            .FontHeader = head
            .HorizontalHeader = hhea
            .HorizontalMetrics = hmtx
            .ANSICharacterMap = cmap
            .Kerning = kern
            .PostScript = post
            .OS2 = os2
            .EmbeddedASCIIHexFontProgram = CreateEmbeddedASCIIHexFontProgram(Bytes, ByteCount)
        End With
        Return x
    End Function
End Module
