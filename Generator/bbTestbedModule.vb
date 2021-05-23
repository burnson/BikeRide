Public Module bbTestbedModule
    Public NotationFont As bbFont
    Public GillSansRegular As bbFont
    Public GillSansItalic As bbFont
    Public FontinSmallCaps As bbFont

    Public Sub MakeTestPDF(ByRef TestFilename As String)
        TotalTime.StartClock()

        Globals.SetDefaults()

        Dim PDF As New bbPDFWriter()

        NotationFont = New bbFont(PDF, 1, "Belle Bonne Revival", "Fonts/BelleBonneRevival.ttf", 8)
        GillSansRegular = New bbFont(PDF, 2, "Gill Sans", "Fonts/GillSans.ttf", 1)
        GillSansItalic = New bbFont(PDF, 3, "Gill Sans Italic", "Fonts/GillSansItalic.ttf", 1)
        FontinSmallCaps = New bbFont(PDF, 4, "Fontin SmallCaps", "Fonts/FontinSmallCaps.ttf", 1)

        PDF.Create(TestFilename)

        Dim Page1 As bbPDFPage = PDF.AddPage(New bbPDFPage(32, 19))

        DrawBikeFrame(Page1)

        Dim TitleText As New bbText(FontinSmallCaps, 0.07)
        Dim SubTitleText As New bbText(FontinSmallCaps, 0.04)
        Page1.AddLineToStream(TitleText.Draw("Bike Ride", New bbRay(0, 0.02 * -9.5, 0), bbText.bbJustification.bbJustifyCenter))
        Page1.AddLineToStream(SubTitleText.Draw("Summer 2007 - Lewisburg, Pennsylvania", New bbRay(0, 0.02 * -13, 0), bbText.bbJustification.bbJustifyCenter))
        Page1.AddLineToStream(TitleText.Draw("Andi Brae", New bbRay(0, 0.02 * -21, 0), bbText.bbJustification.bbJustifyCenter))

        PDF.Close()
        TotalTime.StopClock()
        PrintAllTimers()
    End Sub
End Module