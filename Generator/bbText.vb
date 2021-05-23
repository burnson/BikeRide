Public Class bbText
    Private internalFont As bbFont
    Private internalunitsFontHeight As Double
    Private internalscaleKerning As Double

    Public Enum bbJustification As Integer
        bbJustifyLeft = -1
        bbJustifyCenter = 0
        bbJustifyRight = 1
    End Enum

    Public Sub New(ByRef Font As bbFont, ByVal unitsFontHeight As Double)
        Me.New(Font, unitsFontHeight, 1.0)
    End Sub

    Public Sub New(ByRef Font As bbFont, ByVal unitsFontHeight As Double, ByVal scaleKerning As Double)
        internalFont = Font
        internalunitsFontHeight = unitsFontHeight * internalFont.Scale()
        internalscaleKerning = scaleKerning
    End Sub

    Public Sub DrawBoxed(ByRef Staff As bbStaff, ByRef Text As String, ByVal unitsX As Double, ByVal spacesY As Double, ByVal LineWidth As Double, ByVal RaisePercentage As Double)
        DrawBoxed(Staff, Text, unitsX, spacesY, LineWidth, RaisePercentage, 0.0)
    End Sub

    Public Sub DrawBoxed(ByRef Staff As bbStaff, ByRef Text As String, ByVal unitsX As Double, ByVal spacesY As Double, ByVal LineWidth As Double, ByVal RaisePercentage As Double, ByVal FudgePercentage As Double)
        Dim PDFParent As bbPDFPage = Nothing
        Dim TextWidth As Double = Width("W")
        Dim O As bbRay = Nothing
        Dim TextHeight As Double = internalFont.Info.OS2.sTypoAscender * (1 / 1000.0) * internalunitsFontHeight * internalscaleKerning

        Staff.FindPlottingPoint(unitsX, spacesY, O, 0)
        Staff.GetParent(PDFParent)
        Dim HeightDisplacement As Double = TextWidth * -RaisePercentage + TextWidth / 2
        Dim i As bbPoint = FindDisplacement(New bbPoint(O.x, O.y), New bbPoint(O.x + Math.Cos(O.angle), O.y + Math.Sin(O.angle)), HeightDisplacement)
        O.x = i.X - Math.Cos(O.angle) * TextWidth * FudgePercentage
        O.y = i.Y - Math.Sin(O.angle) * TextWidth * FudgePercentage
        PDFParent.AddLineToStream(DrawSquare(O, TextWidth, LineWidth))
        DrawHorizontalToStaff(Staff, Text, unitsX, spacesY, bbJustification.bbJustifyCenter, 0.0)
    End Sub
    Public Sub DrawHorizontalToStaff(ByRef Staff As bbStaff, ByRef Text As String, ByVal unitsX As Double, ByVal spacesY As Double, ByVal Justification As bbJustification)
        DrawHorizontalToStaff(Staff, Text, unitsX, spacesY, Justification, 0.0)
    End Sub

    Public Sub DrawHorizontalToStaff(ByRef Staff As bbStaff, ByRef Text As String, ByVal unitsX As Double, ByVal spacesY As Double, ByVal Justification As bbJustification, ByVal SegmentKerning As Double)
        Dim PDFParent As bbPDFPage = Nothing
        Dim x As String = ""
        Staff.GetParent(PDFParent)

        'Enter the text draw operator
        x &= SaveCTM()
        x &= "BT" & TokenSeperator()
        x &= internalFont.GetPDFID() & CSpace() & CReal(internalunitsFontHeight) & CSpace() & "Tf" & TokenSeperator()

        Dim relativeX As Double = unitsX
        Dim relativeY As Double = spacesY

        'Determine the justification
        Dim TextWidth As Double = Width(Text)
        If Justification = bbJustification.bbJustifyCenter Then
            relativeX -= TextWidth / 2.0
        ElseIf Justification = bbJustification.bbJustifyRight Then
            relativeX -= TextWidth
        End If

        Dim physicalOrientation As bbRay = Nothing

        Dim LG As Integer
        Dim RG As Integer

        Dim Length As Integer = Len(Text)

        Dim NonKernAdvanceWidth As Double
        Dim Kern As Double
        Dim NextAdvanceWidth As Double

        Dim TextIntegerArray() As Byte
        TextIntegerArray = System.Text.Encoding.ASCII.GetBytes(Text)

        With internalFont.Info

            'translate and rotate to the designated spot
            Staff.FindPlottingPoint(relativeX, relativeY, physicalOrientation, 0)
            x &= CTMTranslate(physicalOrientation.x, physicalOrientation.y)
            x &= CTMRotate(physicalOrientation.angle)

            x &= CReal(0) & CSpace() & CReal(0) & CSpace() & "Td" & TokenSeperator()
            x &= "(" & Chr(TextIntegerArray(0)) & ") Tj" & TokenSeperator()

            'perform manual state reverse
            x &= CTMRotate(-physicalOrientation.angle)
            x &= CTMTranslate(-physicalOrientation.x, -physicalOrientation.y)



            For i As Integer = 0 To Length - 2
                LG = .ANSICharacterMap.CharacterToGlyphMap(TextIntegerArray(i))
                RG = .ANSICharacterMap.CharacterToGlyphMap(TextIntegerArray(i + 1))

                Kern = 0.0
                With .Kerning
                    For j As Integer = 0 To .KerningPairCount - 1
                        If .GlyphLeft(j) = LG And .GlyphRight(j) = RG Then
                            Kern = .KernValue(j)
                        End If
                    Next
                End With

                NonKernAdvanceWidth = .HorizontalMetrics.AdvanceWidth(LG)
                NextAdvanceWidth = NonKernAdvanceWidth + Kern
                NextAdvanceWidth *= (1 / 1000.0) * internalunitsFontHeight * internalscaleKerning

                'translate and rotate to the designated spot
                relativeX += NextAdvanceWidth
                Staff.FindPlottingPointEx(relativeX, relativeY, physicalOrientation, 0, 0, 1.0 + SegmentKerning)
                x &= CTMTranslate(physicalOrientation.x, physicalOrientation.y)
                x &= CTMRotate(physicalOrientation.angle)

                'write character
                x &= CReal(0) & CSpace() & CReal(0) & CSpace() & "Td" & TokenSeperator()
                x &= "(" & Chr(TextIntegerArray(i + 1)) & ") Tj" & TokenSeperator()

                'perform manual state reverse
                x &= CTMRotate(-physicalOrientation.angle)
                x &= CTMTranslate(-physicalOrientation.x, -physicalOrientation.y)
            Next i
        End With
        x &= "ET" & TokenSeperator()
        x &= RevertCTM()

        PDFParent.AddLineToStream(x)
    End Sub

    Public Function Draw(ByRef Text As String, ByRef Anchor As bbRay, ByVal Justification As bbJustification) As String
        Dim CurrentX As Double = Anchor.x
        Dim CurrentY As Double = Anchor.y
        Dim CurrentAngle As Double = Anchor.angle

        If Justification = bbJustification.bbJustifyCenter Then
            Dim TextWidth As Double = Width(Text)
            CurrentX -= Math.Cos(Anchor.angle) * TextWidth / 2.0
            CurrentY -= Math.Sin(Anchor.angle) * TextWidth / 2.0
        Else
            Dim TextWidth As Double = Width(Text)
            CurrentX -= Math.Cos(Anchor.angle) * TextWidth
            CurrentY -= Math.Sin(Anchor.angle) * TextWidth
        End If


        Dim x As String = ""
        x &= SaveCTM()
        x &= CTMTranslate(CurrentX, CurrentY)
        x &= CTMRotate(CurrentAngle) & TokenSeperator()
        x &= "BT" & TokenSeperator()
        x &= internalFont.GetPDFID() & CSpace() & CReal(internalunitsFontHeight) & CSpace() & "Tf" & TokenSeperator()

        Dim LG As Integer
        Dim RG As Integer

        Dim Length As Integer = Len(Text)

        Dim NonKernAdvanceWidth As Double
        Dim Kern As Double
        Dim NextAdvanceWidth As Double

        Dim TextIntegerArray() As Byte
        TextIntegerArray = System.Text.Encoding.ASCII.GetBytes(Text)

        x &= CReal(0) & CSpace() & CReal(0) & CSpace() & "Td" & TokenSeperator()
        x &= "(" & Chr(TextIntegerArray(0)) & ") Tj" & TokenSeperator()

        With internalFont.Info
            For i As Integer = 0 To Length - 2
                LG = .ANSICharacterMap.CharacterToGlyphMap(TextIntegerArray(i))
                RG = .ANSICharacterMap.CharacterToGlyphMap(TextIntegerArray(i + 1))

                Kern = 0.0
                With .Kerning
                    For j As Integer = 0 To .KerningPairCount - 1
                        If .GlyphLeft(j) = LG And .GlyphRight(j) = RG Then
                            Kern = .KernValue(j)
                        End If
                    Next
                End With

                NonKernAdvanceWidth = .HorizontalMetrics.AdvanceWidth(LG)
                NextAdvanceWidth = NonKernAdvanceWidth + Kern
                NextAdvanceWidth /= 1000.0 'Scale by font units
                NextAdvanceWidth *= internalunitsFontHeight * internalscaleKerning

                x &= CReal(NextAdvanceWidth) & CSpace() & CReal(0) & CSpace() & "Td" & TokenSeperator()
                x &= "(" & Chr(TextIntegerArray(i + 1)) & ") Tj" & TokenSeperator()
            Next i
        End With
        x &= "ET" & TokenSeperator()
        x &= RevertCTM()
        Return x
    End Function

    Public Function Width(ByRef Text As String) As Double
        Dim LG As Integer
        Dim RG As Integer

        Dim Length As Integer = Len(Text)

        Dim NonKernAdvanceWidth As Double
        Dim Kern As Double
        Dim NextAdvanceWidth As Double

        Dim TotalWidth As Double = 0

        Dim TextIntegerArray() As Byte
        TextIntegerArray = System.Text.Encoding.ASCII.GetBytes(Text)

        With internalFont.Info
            For i As Integer = 0 To Length - 1
                LG = .ANSICharacterMap.CharacterToGlyphMap(TextIntegerArray(i))
                Kern = 0.0
                If i < Length - 1 Then
                    RG = .ANSICharacterMap.CharacterToGlyphMap(TextIntegerArray(i + 1))
                    With .Kerning
                        For j As Integer = 0 To .KerningPairCount - 1
                            If .GlyphLeft(j) = LG And .GlyphRight(j) = RG Then
                                Kern = .KernValue(j)
                            End If
                        Next
                    End With
                End If
                NonKernAdvanceWidth = .HorizontalMetrics.AdvanceWidth(LG)
                NextAdvanceWidth = NonKernAdvanceWidth + Kern
                NextAdvanceWidth /= 1000.0 'Scale by font units
                NextAdvanceWidth *= internalunitsFontHeight * internalscaleKerning
                TotalWidth += NextAdvanceWidth
            Next i
        End With
        Return TotalWidth
    End Function
End Class
