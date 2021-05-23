Public Class bbPDFWriter
    Private PDFStream As System.IO.FileStream
    Private PDFObjects As New Collection
    Private PDFFonts As New Collection
    Private PDFPages As New Collection

    Public Function GetPDFObjectsCollection() As Collection
        Return PDFObjects
    End Function

    Public Function GetPDFFontsCollection() As Collection
        Return PDFFonts
    End Function

    Private Sub WriteLine(ByRef str As String)
        'Function interprets string as byte array and writes it to the file
        Dim str_Length As Integer = Len(Str)
        Dim byteArray(str_Length + 1) As Byte
        For i As Integer = 0 To str_Length - 1
            byteArray(i) = CByte(Asc(Mid(Str, i + 1, 1)) Mod 256)
        Next
        'The following adds the return characters
        byteArray(str_Length) = 13
        byteArray(str_Length + 1) = 10
        'Write it to the stream...
        PDFStream.Write(byteArray, 0, str_Length + 2)
    End Sub

    Public Function Create(ByRef Filename As String) As Boolean
        'Create a unique filename if the output file already exists
        Dim FinalFilename As String = Filename
        Do While System.IO.File.Exists(FinalFilename)
            FinalFilename = Mid(Filename, 1, Len(Filename) - 4) & "-" & Format(Rnd() * 1000, "0000") & ".pdf"
        Loop
        Filename = FinalFilename

        'Open up the stream
        PDFStream = New System.IO.FileStream(Filename, IO.FileMode.CreateNew, IO.FileAccess.ReadWrite)

        'Write the file type tag
        WriteLine("%PDF-1.3")

        'Write the header object
        Dim Header As New bbPDFObject '1
        Header.Write("<< /Type /Catalog")
        Header.Write("/Pages " & PDFObjectStringID(PDFFonts.Count() * 2 + 2))
        Header.Write(">>")
        PDFObjects.Add(Header)

        'Return success
        Return True
    End Function

    Public Function AddPage(ByRef Page As bbPDFPage) As bbPDFPage
        PDFPages.Add(Page)
        Return Page
    End Function

    Public Function Close() As Boolean
        'All of the PDF file is stored in memory at this point so it needs to be
        'written to file in an organized manner

        'Add this first so its in the right order
        Dim PageInfo As New bbPDFObject '2+FontCount*2 0 R
        PDFObjects.Add(PageInfo)

        Dim PageObjectIndex As Integer = 3
        For i As Integer = 1 To PDFPages.Count
            Dim CurrentPage As bbPDFPage = PDFPages.Item(i)
            With CurrentPage
                Dim Page1Layout As New bbPDFObject '3
                Page1Layout.Write("<< /Type /Page")
                Page1Layout.Write("/Parent " & PDFObjectStringID(2 + PDFFonts.Count() * 2)) 'always 2 + font count * 2
                Page1Layout.Write("/MediaBox [0 0 " & CStr(CurrentPage.Width * 72) & " " & CStr(CurrentPage.Height * 72) & "]")
                Page1Layout.Write("/Contents " & PDFObjectStringID(PageObjectIndex + 1 + PDFFonts.Count() * 2))
                'Begin resources dictionary
                Page1Layout.Write("/Resources <<")
                Page1Layout.Write("/ProcSet [" & PDFObjectStringID(PageObjectIndex + 2 + PDFFonts.Count() * 2) & "]")

                'Load all of the fonts into the dictionary
                If PDFFonts.Count > 0 Then
                    Page1Layout.Write("/Font <<") 'begin fonts

                    For FontIndex As Integer = 1 To PDFFonts.Count()
                        Dim FontObject As bbFont = PDFFonts.Item(FontIndex)
                        Page1Layout.Write(FontObject.GetPDFID() & CSpace() & FontObject.GetPDFObjectID())
                    Next
                    Page1Layout.Write(">>") 'end fonts
                End If

                'End resources dictionary
                Page1Layout.Write(">> %end resources dictionary")

                'Page1Layout.Write("/Font <</F1 x 0 R >>") 'Page1Layout.Write(">>")
                Page1Layout.Write(">>")
                PDFObjects.Add(Page1Layout)


                Dim Page1Content As New bbPDFObject '4
                Page1Content.Write("<< /Length " & .StreamLength & ">>")
                Page1Content.Write("stream")
                Page1Content.Write(.Stream)
                Page1Content.Write("endstream")
                PDFObjects.Add(Page1Content)
            End With

            Dim Page1End As New bbPDFObject '5
            Page1End.Write("[/PDF]")
            PDFObjects.Add(Page1End)

            PageObjectIndex += 3

        Next

        'Now that we know how many pages there are we can write the
        'cataloging information to the PDF object
        PageInfo.Write("<< /Type /Pages")
        'Take care of the "kids" array
        Dim KidsArray As String = ""
        For i As Integer = 0 To PDFPages.Count - 1
            KidsArray &= PDFObjectStringID(3 * i + 3 + PDFFonts.Count() * 2)
        Next
        PageInfo.Write("/Kids [" & KidsArray & "]")
        PageInfo.Write("/Count " & CStr(PDFPages.Count))
        PageInfo.Write(">>")

        'Now that the objects have been initialized, write them all to file
        Dim ObjectList(PDFObjects.Count) As Integer
        For i As Integer = 1 To PDFObjects.Count
            'Save the position of the object in the file for the TOC
            ObjectList(i - 1) = PDFStream.Position()
            'Write the index of the object
            WriteLine(CStr(i) & " 0 obj")

            'Get the code inside the object and write it to the stream
            Dim CurrentObject As bbPDFObject = PDFObjects.Item(i)
            For j As Integer = 1 To CurrentObject.LineCount
                WriteLine(CurrentObject.GetLine(j))
            Next

            'End writing the object
            WriteLine("endobj")
        Next

        'Write the TOC
        ObjectList(PDFObjects.Count) = PDFStream.Position() 'the last position reference goes to xref
        WriteLine("xref")
        WriteLine("0 " & CStr(PDFObjects.Count + 1))
        WriteLine("0000000000 65535 f")
        For i As Integer = 0 To PDFObjects.Count - 1
            WriteLine(Format(ObjectList(i), "0000000000") & " " & "00000 n")
        Next

        WriteLine("trailer")
        WriteLine("<< /Size " & CStr(PDFObjects.Count + 1))
        WriteLine("/Root " & PDFObjectStringID(PDFFonts.Count * 2 + 1))
        WriteLine(">>")
        WriteLine("startxref")
        WriteLine(CStr(ObjectList(PDFObjects.Count())))
        WriteLine("%%EOF")

        'Flush any remaining information to the file and close it
        PDFStream.Flush()
        PDFStream.Close()
    End Function
End Class