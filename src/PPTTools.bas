Attribute VB_Name = "PPTTools"
Option Explicit

Private Const PROGRESS_PREFIX As String = "__ppt_tools_progress_"
Private Const IGNORED_SECTION_NAME As String = "Backup"
Private Const HEADER_HEIGHT As Single = 24
Private Const PAGE_NUMBER_WIDTH As Single = 24

Private Const CODE_EDITOR_SCRIPT As String = "PPTToolsCodePrettier.scpt"
Private Const CODE_FONT_NAME As String = "Consolas"
Private Const CODE_TAG As String = "PPT_TOOLS_CODE"
Private Const CODE_SOURCE_TAG As String = "PPT_TOOLS_CODE_SOURCE"
Private Const CODE_LANGUAGE_TAG As String = "PPT_TOOLS_CODE_LANGUAGE"
Private Const CODE_FONT_SIZE_TAG As String = "PPT_TOOLS_CODE_FONT_SIZE"
Private Const CODE_LINE_NUMBERS_TAG As String = "PPT_TOOLS_CODE_LINE_NUMBERS"
Private Const CODE_NUMBER_GAP_TAG As String = "PPT_TOOLS_CODE_NUMBER_GAP"
Private Const CODE_LINE_HIGHLIGHTS_TAG As String = "PPT_TOOLS_CODE_LINE_HIGHLIGHTS"

Public Sub ProgressBars_Update()
    Dim pres As Presentation
    Dim sld As Slide
    Dim nativeSectionCount As Long
    Dim sourceSectionCount As Long
    Dim displayedSectionCount As Long
    Dim displayedSectionIndex As Long
    Dim currentSection As Long
    Dim currentPageNumber As Long
    Dim sectionNumber As Long
    Dim sectionName As String
    Dim sectionFirstSlide As Long
    Dim sectionSlideCount As Long
    Dim sectionProgress As Single
    Dim slideWidth As Single
    Dim barWidth As Single
    Dim cellWidth As Single
    Dim segmentLeft As Single
    Dim segmentWidth As Single

    On Error GoTo UpdateFailed

    Set pres = Application.ActivePresentation
    If pres.Slides.Count = 0 Then Exit Sub

    nativeSectionCount = pres.SectionProperties.Count
    If nativeSectionCount > 0 Then
        sourceSectionCount = nativeSectionCount
        displayedSectionCount = GetDisplayedSectionCount(pres)
    Else
        sourceSectionCount = 1
        displayedSectionCount = 1
    End If

    If displayedSectionCount = 0 Then
        For Each sld In pres.Slides
            DeleteProgressShapes sld
        Next sld

        MsgBox "Section progress updated.", vbInformation, "PPT Tools"
        Exit Sub
    End If

    slideWidth = pres.PageSetup.SlideWidth
    barWidth = slideWidth - PAGE_NUMBER_WIDTH
    cellWidth = barWidth / displayedSectionCount

    For Each sld In pres.Slides
        DeleteProgressShapes sld

        If nativeSectionCount > 0 Then
            currentSection = CLng(sld.SectionIndex)
        Else
            currentSection = 1
        End If

        If nativeSectionCount = 0 Or _
            Not IsIgnoredSection(pres, currentSection) Then

            currentPageNumber = currentPageNumber + 1
            displayedSectionIndex = 0

            For sectionNumber = 1 To sourceSectionCount
                If nativeSectionCount = 0 Or _
                    Not IsIgnoredSection(pres, sectionNumber) Then

                    displayedSectionIndex = displayedSectionIndex + 1
                    segmentLeft = _
                        (displayedSectionIndex - 1) * cellWidth
                    segmentWidth = cellWidth

                    If nativeSectionCount > 0 Then
                        sectionName = _
                            pres.SectionProperties.Name(sectionNumber)
                        sectionFirstSlide = _
                            pres.SectionProperties.FirstSlide(sectionNumber)
                        sectionSlideCount = _
                            pres.SectionProperties.SlidesCount(sectionNumber)
                    Else
                        sectionName = "Presentation"
                        sectionFirstSlide = 1
                        sectionSlideCount = pres.Slides.Count
                    End If

                    sectionProgress = GetSectionProgress( _
                        sld.SlideIndex, _
                        sectionNumber, _
                        currentSection, _
                        sectionFirstSlide, _
                        sectionSlideCount)

                    AddSegmentBackground sld, displayedSectionIndex, _
                        segmentLeft, segmentWidth

                    If sectionProgress > 0 Then
                        AddSegmentFill sld, displayedSectionIndex, _
                            segmentLeft, segmentWidth * sectionProgress
                    End If

                    If sectionNumber = currentSection Then
                        AddPositionMarker sld, _
                            segmentLeft + _
                            (segmentWidth * sectionProgress)
                    End If

                    AddSectionLabel sld, displayedSectionIndex, _
                        sectionName, segmentLeft, segmentWidth, _
                        sectionNumber = currentSection

                End If
            Next sectionNumber

            For displayedSectionIndex = 1 To _
                displayedSectionCount - 1
                AddSectionDivider sld, displayedSectionIndex, _
                    displayedSectionIndex * cellWidth
            Next displayedSectionIndex

            AddPageNumber sld, currentPageNumber, barWidth
        End If
    Next sld

    MsgBox "Section progress updated.", vbInformation, "PPT Tools"
    Exit Sub

UpdateFailed:
    MsgBox "Unable to update section progress: " & Err.Description, _
        vbExclamation, "PPT Tools"
End Sub

Public Sub ProgressBars_Remove()
    Dim pres As Presentation
    Dim sld As Slide

    On Error GoTo RemoveFailed

    Set pres = Application.ActivePresentation

    For Each sld In pres.Slides
        DeleteProgressShapes sld
    Next sld

    MsgBox "Section progress removed.", vbInformation, "PPT Tools"
    Exit Sub

RemoveFailed:
    MsgBox "Unable to remove section progress: " & Err.Description, _
        vbExclamation, "PPT Tools"
End Sub

Public Sub RibbonUpdateProgressBars(ByVal control As IRibbonControl)
    ProgressBars_Update
End Sub

Public Sub RibbonRemoveProgressBars(ByVal control As IRibbonControl)
    ProgressBars_Remove
End Sub

Public Sub Notes_ExportMarkdown()
    Dim pres As Presentation
    Dim selectedFolder As String
    Dim outputPath As String
    Dim markdownContent As String
    Dim notesText As String
    Dim sld As Slide

    On Error GoTo ExportFailed

    Set pres = Application.ActivePresentation
    If pres.Slides.Count = 0 Then
        MsgBox "The active presentation has no slides.", _
            vbInformation, "PPT Tools"
        Exit Sub
    End If

    On Error Resume Next
    selectedFolder = MacScript( _
        "return POSIX path of (choose folder with prompt """ & _
        "Choose a folder for the Markdown file""" & _
        ")")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo ExportFailed
        Exit Sub
    End If
    On Error GoTo ExportFailed

    If Len(selectedFolder) = 0 Then Exit Sub

    outputPath = GetAvailableNotesPath(selectedFolder, pres.Name)

    For Each sld In pres.Slides
        markdownContent = markdownContent & _
            "# Slide " & CStr(sld.SlideIndex) & vbLf & vbLf

        notesText = GetSlideNotes(sld)
        If Len(notesText) > 0 Then
            markdownContent = markdownContent & notesText & vbLf & vbLf
        End If
    Next sld

    WriteUtf8File outputPath, markdownContent

    MsgBox "Notes exported to:" & vbCrLf & outputPath, _
        vbInformation, "PPT Tools"
    Exit Sub

ExportFailed:
    MsgBox "Unable to export notes: " & Err.Description, _
        vbExclamation, "PPT Tools"
End Sub

Public Sub RibbonExportNotesToMarkdown(ByVal control As IRibbonControl)
    Notes_ExportMarkdown
End Sub

Public Sub Notes_LoadMarkdown()
    Dim pres As Presentation
    Dim selectedFile As String
    Dim markdownContent As String
    Dim slideNumbers() As Long
    Dim noteBlocks() As String
    Dim blockCount As Long
    Dim blockIndex As Long
    Dim loadedCount As Long
    Dim skippedCount As Long

    On Error GoTo LoadFailed

    Set pres = Application.ActivePresentation
    If pres.Slides.Count = 0 Then
        MsgBox "The active presentation has no slides.", _
            vbInformation, "PPT Tools"
        Exit Sub
    End If

    On Error Resume Next
    selectedFile = MacScript( _
        "return POSIX path of (choose file with prompt """ & _
        "Choose a Markdown notes file""" & _
        ")")
    If Err.Number <> 0 Then
        Err.Clear
        On Error GoTo LoadFailed
        Exit Sub
    End If
    On Error GoTo LoadFailed

    If Len(selectedFile) = 0 Then Exit Sub
    If Not IsMarkdownPath(selectedFile) Then
        MsgBox "Choose a .md or .markdown file.", _
            vbInformation, "PPT Tools"
        Exit Sub
    End If

    markdownContent = ReadUtf8File(selectedFile)
    ParseMarkdownNotes markdownContent, _
        slideNumbers, noteBlocks, blockCount

    If blockCount = 0 Then
        MsgBox "No # Slide N headings were found.", _
            vbInformation, "PPT Tools"
        Exit Sub
    End If

    For blockIndex = 1 To blockCount
        If slideNumbers(blockIndex) >= 1 And _
            slideNumbers(blockIndex) <= pres.Slides.Count Then

            If SetSlideNotes( _
                pres.Slides(slideNumbers(blockIndex)), _
                noteBlocks(blockIndex)) Then
                loadedCount = loadedCount + 1
            Else
                skippedCount = skippedCount + 1
            End If
        Else
            skippedCount = skippedCount + 1
        End If
    Next blockIndex

    MsgBox "Notes loaded for " & CStr(loadedCount) & " slide(s)." & _
        IIf(skippedCount > 0, _
            vbCrLf & CStr(skippedCount) & " block(s) skipped.", _
            ""), _
        vbInformation, "PPT Tools"
    Exit Sub

LoadFailed:
    MsgBox "Unable to load notes: " & Err.Description, _
        vbExclamation, "PPT Tools"
End Sub

Public Sub RibbonLoadNotesFromMarkdown(ByVal control As IRibbonControl)
    Notes_LoadMarkdown
End Sub

Public Sub CodePrettier_Open()
    Dim targetSlide As Slide
    Dim codeShape As Shape
    Dim resultText As String

    On Error GoTo OpenFailed

    Set codeShape = GetSelectedCodeShape()
    If codeShape Is Nothing Then
        Set targetSlide = GetCurrentSlide()
        If targetSlide Is Nothing Then
            MsgBox "Open a slide in Normal view first.", _
                vbInformation, "PPT Tools"
            Exit Sub
        End If
    End If

    On Error Resume Next
    resultText = AppleScriptTask( _
        CODE_EDITOR_SCRIPT, "openEditor", BuildCodeEditorState(codeShape))
    If Err.Number <> 0 Then
        MsgBox "Unable to open the Code Prettier editor: " & _
            Err.Description & vbCrLf & vbCrLf & _
            "Run scripts/install_code_prettier.sh to install the editor.", _
            vbExclamation, "PPT Tools"
        Err.Clear
        Exit Sub
    End If
    On Error GoTo OpenFailed

    ApplyCodeEditorResult targetSlide, codeShape, resultText
    Exit Sub

OpenFailed:
    MsgBox "Unable to update the code block: " & Err.Description, _
        vbExclamation, "PPT Tools"
End Sub

Public Sub RibbonCodePrettier(ByVal control As IRibbonControl)
    CodePrettier_Open
End Sub

Private Function GetDisplayedSectionCount( _
    ByVal pres As Presentation) As Long

    Dim sectionNumber As Long

    For sectionNumber = 1 To pres.SectionProperties.Count
        If Not IsIgnoredSection(pres, sectionNumber) Then
            GetDisplayedSectionCount = GetDisplayedSectionCount + 1
        End If
    Next sectionNumber
End Function

Private Function IsIgnoredSection( _
    ByVal pres As Presentation, _
    ByVal sectionNumber As Long) As Boolean

    IsIgnoredSection = StrComp( _
        Trim$(pres.SectionProperties.Name(sectionNumber)), _
        IGNORED_SECTION_NAME, _
        vbTextCompare) = 0
End Function

Private Function GetSectionProgress( _
    ByVal slideIndex As Long, _
    ByVal sectionNumber As Long, _
    ByVal currentSection As Long, _
    ByVal sectionFirstSlide As Long, _
    ByVal sectionSlideCount As Long) As Single

    If sectionNumber < currentSection Then
        GetSectionProgress = 1
    ElseIf sectionNumber > currentSection Or sectionSlideCount = 0 Then
        GetSectionProgress = 0
    Else
        GetSectionProgress = _
            (slideIndex - sectionFirstSlide + 1) / sectionSlideCount
    End If
End Function

Private Sub AddSegmentBackground( _
    ByVal sld As Slide, _
    ByVal sectionNumber As Long, _
    ByVal leftPos As Single, _
    ByVal segmentWidth As Single)

    AddRectangle sld, _
        PROGRESS_PREFIX & "track_" & CStr(sectionNumber), _
        leftPos, segmentWidth, RGB(229, 233, 239)
End Sub

Private Sub AddSegmentFill( _
    ByVal sld As Slide, _
    ByVal sectionNumber As Long, _
    ByVal leftPos As Single, _
    ByVal fillWidth As Single)

    AddRectangle sld, _
        PROGRESS_PREFIX & "fill_" & CStr(sectionNumber), _
        leftPos, fillWidth, RGB(183, 202, 238)
End Sub

Private Sub AddRectangle( _
    ByVal sld As Slide, _
    ByVal shapeName As String, _
    ByVal leftPos As Single, _
    ByVal shapeWidth As Single, _
    ByVal fillColor As Long)

    Dim shp As Shape

    Set shp = sld.Shapes.AddShape( _
        msoShapeRectangle, leftPos, 0, shapeWidth, HEADER_HEIGHT)

    With shp
        .Name = shapeName
        .AlternativeText = "PPT Tools section progress"
        .Line.Visible = msoFalse
        .Fill.Solid
        .Fill.ForeColor.RGB = fillColor
    End With
End Sub

Private Sub AddSectionDivider( _
    ByVal sld As Slide, _
    ByVal sectionNumber As Long, _
    ByVal leftPos As Single)

    Dim divider As Shape

    Set divider = sld.Shapes.AddShape( _
        msoShapeRectangle, _
        leftPos - 1, 0, 2, HEADER_HEIGHT)

    With divider
        .Name = PROGRESS_PREFIX & "divider_" & CStr(sectionNumber)
        .AlternativeText = "PPT Tools section divider"
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(0, 0, 0)
        .Line.Visible = msoFalse
    End With
End Sub

Private Sub AddPositionMarker( _
    ByVal sld As Slide, _
    ByVal leftPos As Single)

    Dim marker As Shape

    Set marker = sld.Shapes.AddLine( _
        leftPos, 0, leftPos, HEADER_HEIGHT)

    With marker
        .Name = PROGRESS_PREFIX & "current_position"
        .AlternativeText = "PPT Tools current position"
        .Line.ForeColor.RGB = RGB(35, 78, 170)
        .Line.Weight = 1.25
    End With
End Sub

Private Sub AddSectionLabel( _
    ByVal sld As Slide, _
    ByVal sectionNumber As Long, _
    ByVal labelText As String, _
    ByVal leftPos As Single, _
    ByVal labelWidth As Single, _
    ByVal isCurrent As Boolean)

    Dim label As Shape

    Set label = sld.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        leftPos, 0, labelWidth, HEADER_HEIGHT)

    With label
        .Name = PROGRESS_PREFIX & "label_" & CStr(sectionNumber)
        .AlternativeText = "PPT Tools section name"
        .Fill.Visible = msoFalse
        .Line.Visible = msoFalse

        With .TextFrame
            .AutoSize = ppAutoSizeNone
            .MarginLeft = 4
            .MarginRight = 4
            .MarginTop = 1
            .MarginBottom = 1
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse

            With .TextRange
                .Text = labelText
                .ParagraphFormat.Alignment = ppAlignCenter
                .Font.Name = "Aptos"
                .Font.Size = 10.5

                If isCurrent Then
                    .Font.Bold = msoTrue
                    .Font.Color.RGB = RGB(35, 78, 170)
                Else
                    .Font.Bold = msoFalse
                    .Font.Color.RGB = RGB(42, 50, 64)
                End If
            End With
        End With

    End With
End Sub

Private Sub AddPageNumber( _
    ByVal sld As Slide, _
    ByVal currentPage As Long, _
    ByVal leftPos As Single)

    Dim pageNumber As Shape

    Set pageNumber = sld.Shapes.AddTextbox( _
        msoTextOrientationHorizontal, _
        leftPos, 0, PAGE_NUMBER_WIDTH, HEADER_HEIGHT)

    With pageNumber
        .Name = PROGRESS_PREFIX & "page_number"
        .AlternativeText = "PPT Tools page number"
        .Fill.Visible = msoTrue
        .Fill.Solid
        .Fill.ForeColor.RGB = RGB(255, 255, 255)
        .Line.Visible = msoTrue
        .Line.ForeColor.RGB = RGB(105, 115, 130)
        .Line.Weight = 0.75

        With .TextFrame
            .AutoSize = ppAutoSizeNone
            .MarginLeft = 0
            .MarginRight = 0
            .MarginTop = 1
            .MarginBottom = 1
            .VerticalAnchor = msoAnchorMiddle
            .WordWrap = msoFalse

            With .TextRange
                .Text = CStr(currentPage)
                .ParagraphFormat.Alignment = ppAlignCenter
                .Font.Name = "Aptos"
                .Font.Size = 10.5
                .Font.Bold = msoFalse
                .Font.Color.RGB = RGB(42, 50, 64)
            End With
        End With
    End With
End Sub

Private Sub DeleteProgressShapes(ByVal sld As Slide)
    Dim shapeIndex As Long
    Dim shapeName As String

    For shapeIndex = sld.Shapes.Count To 1 Step -1
        shapeName = sld.Shapes(shapeIndex).Name
        If Left$(shapeName, Len(PROGRESS_PREFIX)) = PROGRESS_PREFIX Then
            sld.Shapes(shapeIndex).Delete
        End If
    Next shapeIndex
End Sub

Private Function GetSlideNotes(ByVal sld As Slide) As String
    Dim placeholder As Shape

    For Each placeholder In sld.NotesPage.Shapes.Placeholders
        If placeholder.PlaceholderFormat.Type = ppPlaceholderBody Then
            If placeholder.HasTextFrame = msoTrue Then
                If placeholder.TextFrame.HasText = msoTrue Then
                    GetSlideNotes = NormalizeNotesText( _
                        placeholder.TextFrame.TextRange.Text)
                End If
            End If
            Exit Function
        End If
    Next placeholder
End Function

Private Function SetSlideNotes( _
    ByVal sld As Slide, _
    ByVal notesText As String) As Boolean

    Dim placeholder As Shape
    Dim notesRange As TextRange
    Dim plainText As String
    Dim boldStarts() As Long
    Dim boldLengths() As Long
    Dim boldCount As Long
    Dim boldIndex As Long

    ParseMarkdownBold notesText, plainText, _
        boldStarts, boldLengths, boldCount

    For Each placeholder In sld.NotesPage.Shapes.Placeholders
        If placeholder.PlaceholderFormat.Type = ppPlaceholderBody Then
            Set notesRange = placeholder.TextFrame.TextRange
            notesRange.Text = Replace(plainText, vbLf, vbCr)

            If Len(plainText) > 0 Then
                notesRange.Font.Bold = msoFalse

                For boldIndex = 1 To boldCount
                    notesRange.Characters( _
                        boldStarts(boldIndex), _
                        boldLengths(boldIndex)).Font.Bold = msoTrue
                Next boldIndex
            End If

            SetSlideNotes = True
            Exit Function
        End If
    Next placeholder
End Function

Private Sub ParseMarkdownBold( _
    ByVal markdownText As String, _
    ByRef plainText As String, _
    ByRef boldStarts() As Long, _
    ByRef boldLengths() As Long, _
    ByRef boldCount As Long)

    Dim scanPosition As Long
    Dim openingPosition As Long
    Dim closingPosition As Long
    Dim boldText As String

    scanPosition = 1

    Do While scanPosition <= Len(markdownText)
        openingPosition = InStr( _
            scanPosition, markdownText, "**", vbBinaryCompare)

        If openingPosition = 0 Then
            plainText = plainText & Mid$(markdownText, scanPosition)
            Exit Do
        End If

        closingPosition = InStr( _
            openingPosition + 2, markdownText, "**", vbBinaryCompare)

        If closingPosition = 0 Then
            plainText = plainText & Mid$(markdownText, scanPosition)
            Exit Do
        End If

        plainText = plainText & Mid$( _
            markdownText, _
            scanPosition, _
            openingPosition - scanPosition)

        boldText = Mid$( _
            markdownText, _
            openingPosition + 2, _
            closingPosition - openingPosition - 2)

        If Len(boldText) > 0 Then
            AppendBoldRange boldStarts, boldLengths, boldCount, _
                Len(plainText) + 1, Len(boldText)
            plainText = plainText & boldText
        End If

        scanPosition = closingPosition + 2
    Loop
End Sub

Private Sub AppendBoldRange( _
    ByRef boldStarts() As Long, _
    ByRef boldLengths() As Long, _
    ByRef boldCount As Long, _
    ByVal rangeStart As Long, _
    ByVal rangeLength As Long)

    boldCount = boldCount + 1

    If boldCount = 1 Then
        ReDim boldStarts(1 To 1)
        ReDim boldLengths(1 To 1)
    Else
        ReDim Preserve boldStarts(1 To boldCount)
        ReDim Preserve boldLengths(1 To boldCount)
    End If

    boldStarts(boldCount) = rangeStart
    boldLengths(boldCount) = rangeLength
End Sub

Private Function NormalizeNotesText(ByVal value As String) As String
    value = Replace(value, vbCrLf, vbLf)
    value = Replace(value, vbCr, vbLf)
    value = Replace(value, Chr$(11), vbLf)

    Do While Len(value) > 0 And Left$(value, 1) = vbLf
        value = Mid$(value, 2)
    Loop

    Do While Len(value) > 0 And Right$(value, 1) = vbLf
        value = Left$(value, Len(value) - 1)
    Loop

    NormalizeNotesText = value
End Function

Private Function IsMarkdownPath(ByVal filePath As String) As Boolean
    Dim lowerPath As String

    lowerPath = LCase$(filePath)
    IsMarkdownPath = _
        Right$(lowerPath, 3) = ".md" Or _
        Right$(lowerPath, 9) = ".markdown"
End Function

Private Sub ParseMarkdownNotes( _
    ByVal markdownContent As String, _
    ByRef slideNumbers() As Long, _
    ByRef noteBlocks() As String, _
    ByRef blockCount As Long)

    Dim lines As Variant
    Dim lineIndex As Long
    Dim currentSlideNumber As Long
    Dim parsedSlideNumber As Long
    Dim currentNotes As String
    Dim hasCurrentBlock As Boolean
    Dim currentLine As String

    markdownContent = Replace(markdownContent, vbCrLf, vbLf)
    markdownContent = Replace(markdownContent, vbCr, vbLf)
    lines = Split(markdownContent, vbLf)

    For lineIndex = LBound(lines) To UBound(lines)
        currentLine = CStr(lines(lineIndex))

        If TryParseSlideHeader(currentLine, parsedSlideNumber) Then
            If hasCurrentBlock Then
                AppendNotesBlock slideNumbers, noteBlocks, blockCount, _
                    currentSlideNumber, currentNotes
            End If

            currentSlideNumber = parsedSlideNumber
            currentNotes = ""
            hasCurrentBlock = True
        ElseIf hasCurrentBlock Then
            If Len(currentNotes) > 0 Then
                currentNotes = currentNotes & vbLf
            End If
            currentNotes = currentNotes & currentLine
        End If
    Next lineIndex

    If hasCurrentBlock Then
        AppendNotesBlock slideNumbers, noteBlocks, blockCount, _
            currentSlideNumber, currentNotes
    End If
End Sub

Private Sub AppendNotesBlock( _
    ByRef slideNumbers() As Long, _
    ByRef noteBlocks() As String, _
    ByRef blockCount As Long, _
    ByVal slideNumber As Long, _
    ByVal notesText As String)

    blockCount = blockCount + 1

    If blockCount = 1 Then
        ReDim slideNumbers(1 To 1)
        ReDim noteBlocks(1 To 1)
    Else
        ReDim Preserve slideNumbers(1 To blockCount)
        ReDim Preserve noteBlocks(1 To blockCount)
    End If

    slideNumbers(blockCount) = slideNumber
    noteBlocks(blockCount) = NormalizeNotesText(notesText)
End Sub

Private Function TryParseSlideHeader( _
    ByVal lineText As String, _
    ByRef slideNumber As Long) As Boolean

    Const HEADER_PREFIX As String = "# Slide "

    Dim trimmedLine As String
    Dim numberText As String
    Dim characterIndex As Long
    Dim characterValue As String

    trimmedLine = Trim$(lineText)
    If Len(trimmedLine) <= Len(HEADER_PREFIX) Then Exit Function
    If StrComp( _
        Left$(trimmedLine, Len(HEADER_PREFIX)), _
        HEADER_PREFIX, _
        vbTextCompare) <> 0 Then Exit Function

    numberText = Trim$(Mid$(trimmedLine, Len(HEADER_PREFIX) + 1))
    If Len(numberText) = 0 Then Exit Function

    For characterIndex = 1 To Len(numberText)
        characterValue = Mid$(numberText, characterIndex, 1)
        If characterValue < "0" Or characterValue > "9" Then
            Exit Function
        End If
    Next characterIndex

    On Error GoTo InvalidHeader
    slideNumber = CLng(numberText)
    If slideNumber < 1 Then Exit Function

    TryParseSlideHeader = True
    Exit Function

InvalidHeader:
    slideNumber = 0
End Function

Private Function GetAvailableNotesPath( _
    ByVal folderPath As String, _
    ByVal presentationName As String) As String

    Dim baseName As String
    Dim basePath As String
    Dim candidatePath As String
    Dim suffix As Long

    baseName = GetPresentationBaseName(presentationName)
    basePath = JoinPath(folderPath, baseName & "-notes")
    candidatePath = basePath & ".md"
    suffix = 2

    Do While Len(Dir$(candidatePath)) > 0
        candidatePath = basePath & "-" & CStr(suffix) & ".md"
        suffix = suffix + 1
    Loop

    GetAvailableNotesPath = candidatePath
End Function

Private Function GetPresentationBaseName( _
    ByVal presentationName As String) As String

    Dim extensionPosition As Long
    Dim invalidCharacter As Variant

    extensionPosition = InStrRev(presentationName, ".")
    If extensionPosition > 1 Then
        presentationName = Left$(presentationName, extensionPosition - 1)
    End If

    For Each invalidCharacter In _
        Array("\", "/", ":", "*", "?", """", "<", ">", "|")
        presentationName = Replace( _
            presentationName, CStr(invalidCharacter), "-")
    Next invalidCharacter

    If Len(presentationName) = 0 Then presentationName = "presentation"
    GetPresentationBaseName = presentationName
End Function

Private Function JoinPath( _
    ByVal folderPath As String, _
    ByVal fileName As String) As String

    Dim separator As String

    If InStr(folderPath, "/") > 0 Then
        separator = "/"
    Else
        separator = "\"
    End If

    If Right$(folderPath, 1) = "/" Or _
        Right$(folderPath, 1) = "\" Then
        JoinPath = folderPath & fileName
    Else
        JoinPath = folderPath & separator & fileName
    End If
End Function

Private Sub WriteUtf8File( _
    ByVal outputPath As String, _
    ByVal contents As String)

    Dim fileNumber As Integer
    Dim encodedBytes() As Byte
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo WriteFailed

    fileNumber = FreeFile
    Open outputPath For Binary Access Write As #fileNumber

    If Len(contents) > 0 Then
        encodedBytes = EncodeUtf8(contents)
        Put #fileNumber, 1, encodedBytes
    End If

    Close #fileNumber
    Exit Sub

WriteFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0
    Err.Raise errorNumber, "PPT Tools", errorDescription
End Sub

Private Function ReadUtf8File(ByVal filePath As String) As String
    Dim fileNumber As Integer
    Dim fileLength As Long
    Dim encodedBytes() As Byte
    Dim errorNumber As Long
    Dim errorDescription As String

    On Error GoTo ReadFailed

    fileNumber = FreeFile
    Open filePath For Binary Access Read As #fileNumber
    fileLength = LOF(fileNumber)

    If fileLength > 0 Then
        ReDim encodedBytes(0 To fileLength - 1)
        Get #fileNumber, 1, encodedBytes
        ReadUtf8File = DecodeUtf8(encodedBytes)
    End If

    Close #fileNumber
    Exit Function

ReadFailed:
    errorNumber = Err.Number
    errorDescription = Err.Description
    On Error Resume Next
    Close #fileNumber
    On Error GoTo 0
    Err.Raise errorNumber, "PPT Tools", errorDescription
End Function

Private Function EncodeUtf8(ByVal value As String) As Byte()
    Dim encodedBytes() As Byte
    Dim characterIndex As Long
    Dim byteIndex As Long
    Dim characterCount As Long
    Dim codePoint As Long
    Dim nextCodeUnit As Long

    characterCount = Len(value)
    ReDim encodedBytes(0 To (characterCount * 3) - 1)

    characterIndex = 1
    Do While characterIndex <= characterCount
        codePoint = AscW(Mid$(value, characterIndex, 1))
        If codePoint < 0 Then codePoint = codePoint + &H10000

        If codePoint >= &HD800& And codePoint <= &HDBFF& And _
            characterIndex < characterCount Then
            nextCodeUnit = AscW(Mid$(value, characterIndex + 1, 1))
            If nextCodeUnit < 0 Then nextCodeUnit = nextCodeUnit + &H10000

            If nextCodeUnit >= &HDC00& And nextCodeUnit <= &HDFFF& Then
                codePoint = &H10000 + _
                    ((codePoint - &HD800&) * &H400) + _
                    (nextCodeUnit - &HDC00&)
                characterIndex = characterIndex + 1
            Else
                codePoint = &HFFFD&
            End If
        ElseIf codePoint >= &HDC00& And codePoint <= &HDFFF& Then
            codePoint = &HFFFD&
        End If

        If codePoint <= &H7F Then
            encodedBytes(byteIndex) = codePoint
            byteIndex = byteIndex + 1
        ElseIf codePoint <= &H7FF Then
            encodedBytes(byteIndex) = _
                &HC0 Or ((codePoint \ &H40) And &H1F)
            encodedBytes(byteIndex + 1) = _
                &H80 Or (codePoint And &H3F)
            byteIndex = byteIndex + 2
        ElseIf codePoint <= &HFFFF& Then
            encodedBytes(byteIndex) = _
                &HE0 Or ((codePoint \ &H1000) And &HF)
            encodedBytes(byteIndex + 1) = _
                &H80 Or ((codePoint \ &H40) And &H3F)
            encodedBytes(byteIndex + 2) = _
                &H80 Or (codePoint And &H3F)
            byteIndex = byteIndex + 3
        Else
            encodedBytes(byteIndex) = _
                &HF0 Or ((codePoint \ &H40000) And &H7)
            encodedBytes(byteIndex + 1) = _
                &H80 Or ((codePoint \ &H1000) And &H3F)
            encodedBytes(byteIndex + 2) = _
                &H80 Or ((codePoint \ &H40) And &H3F)
            encodedBytes(byteIndex + 3) = _
                &H80 Or (codePoint And &H3F)
            byteIndex = byteIndex + 4
        End If

        characterIndex = characterIndex + 1
    Loop

    ReDim Preserve encodedBytes(0 To byteIndex - 1)
    EncodeUtf8 = encodedBytes
End Function

Private Function DecodeUtf8(ByRef encodedBytes() As Byte) As String
    Dim byteIndex As Long
    Dim lastByteIndex As Long
    Dim firstByte As Long
    Dim codePoint As Long
    Dim sequenceLength As Long
    Dim decodedText As String

    byteIndex = LBound(encodedBytes)
    lastByteIndex = UBound(encodedBytes)

    If lastByteIndex - byteIndex >= 2 Then
        If encodedBytes(byteIndex) = &HEF And _
            encodedBytes(byteIndex + 1) = &HBB And _
            encodedBytes(byteIndex + 2) = &HBF Then
            byteIndex = byteIndex + 3
        End If
    End If

    Do While byteIndex <= lastByteIndex
        firstByte = encodedBytes(byteIndex)
        codePoint = &HFFFD&
        sequenceLength = 1

        If firstByte <= &H7F Then
            codePoint = firstByte
        ElseIf firstByte >= &HC2 And firstByte <= &HDF And _
            byteIndex + 1 <= lastByteIndex Then

            If IsUtf8Continuation(encodedBytes(byteIndex + 1)) Then
                codePoint = _
                    (CLng(firstByte And &H1F) * &H40) + _
                    (encodedBytes(byteIndex + 1) And &H3F)
                sequenceLength = 2
            End If
        ElseIf firstByte >= &HE0 And firstByte <= &HEF And _
            byteIndex + 2 <= lastByteIndex Then

            If IsUtf8Continuation(encodedBytes(byteIndex + 1)) And _
                IsUtf8Continuation(encodedBytes(byteIndex + 2)) Then

                codePoint = _
                    (CLng(firstByte And &HF) * &H1000) + _
                    (CLng(encodedBytes(byteIndex + 1) And &H3F) * &H40) + _
                    (encodedBytes(byteIndex + 2) And &H3F)
                sequenceLength = 3

                If codePoint < &H800 Or _
                    (codePoint >= &HD800& And codePoint <= &HDFFF&) Then
                    codePoint = &HFFFD&
                End If
            End If
        ElseIf firstByte >= &HF0 And firstByte <= &HF4 And _
            byteIndex + 3 <= lastByteIndex Then

            If IsUtf8Continuation(encodedBytes(byteIndex + 1)) And _
                IsUtf8Continuation(encodedBytes(byteIndex + 2)) And _
                IsUtf8Continuation(encodedBytes(byteIndex + 3)) Then

                codePoint = _
                    (CLng(firstByte And &H7) * &H40000) + _
                    (CLng(encodedBytes(byteIndex + 1) And &H3F) * &H1000) + _
                    (CLng(encodedBytes(byteIndex + 2) And &H3F) * &H40) + _
                    (encodedBytes(byteIndex + 3) And &H3F)
                sequenceLength = 4

                If codePoint < &H10000 Or codePoint > &H10FFFF Then
                    codePoint = &HFFFD&
                End If
            End If
        End If

        decodedText = decodedText & CodePointToString(codePoint)
        byteIndex = byteIndex + sequenceLength
    Loop

    DecodeUtf8 = decodedText
End Function

Private Function IsUtf8Continuation(ByVal byteValue As Long) As Boolean
    IsUtf8Continuation = _
        byteValue >= &H80 And byteValue <= &HBF
End Function

Private Function CodePointToString(ByVal codePoint As Long) As String
    Dim highSurrogate As Long
    Dim lowSurrogate As Long

    If codePoint <= &HFFFF& Then
        CodePointToString = ChrWCodeUnit(codePoint)
    Else
        codePoint = codePoint - &H10000
        highSurrogate = &HD800& + (codePoint \ &H400)
        lowSurrogate = &HDC00& + (codePoint And &H3FF)
        CodePointToString = _
            ChrWCodeUnit(highSurrogate) & _
            ChrWCodeUnit(lowSurrogate)
    End If
End Function

Private Function ChrWCodeUnit(ByVal codeUnit As Long) As String
    If codeUnit > &H7FFF& Then codeUnit = codeUnit - &H10000
    ChrWCodeUnit = ChrW(codeUnit)
End Function

Private Function GetCurrentSlide() As Slide
    On Error Resume Next
    Set GetCurrentSlide = Application.ActiveWindow.View.Slide
End Function

Private Function GetSelectedCodeShape() As Shape
    Dim selectedShapes As ShapeRange

    On Error Resume Next
    Set selectedShapes = Application.ActiveWindow.Selection.ShapeRange
    On Error GoTo 0

    If selectedShapes Is Nothing Then Exit Function
    If selectedShapes.Count <> 1 Then Exit Function

    If selectedShapes(1).Tags.Item(CODE_TAG) = "1" Then
        Set GetSelectedCodeShape = selectedShapes(1)
    End If
End Function

' Editor protocol: one "key=value" record per line; values escape
' backslash, line feed, carriage return and tab as \\ \n \r \t.
Private Function BuildCodeEditorState(ByVal codeShape As Shape) As String
    If codeShape Is Nothing Then
        BuildCodeEditorState = "mode=insert"
        Exit Function
    End If

    With codeShape.Tags
        BuildCodeEditorState = "mode=edit" & vbLf & _
            "language=" & EscapeCodeField(.Item(CODE_LANGUAGE_TAG)) & vbLf & _
            "fontSize=" & .Item(CODE_FONT_SIZE_TAG) & vbLf & _
            "lineNumbers=" & .Item(CODE_LINE_NUMBERS_TAG) & vbLf & _
            "numberGap=" & .Item(CODE_NUMBER_GAP_TAG) & vbLf & _
            "lineHighlights=" & .Item(CODE_LINE_HIGHLIGHTS_TAG) & vbLf & _
            "source=" & EscapeCodeField(.Item(CODE_SOURCE_TAG))
    End With
End Function

Private Sub ApplyCodeEditorResult( _
    ByVal targetSlide As Slide, _
    ByVal codeShape As Shape, _
    ByVal resultText As String)

    Dim records As Variant
    Dim recordIndex As Long
    Dim recordText As String
    Dim separatorPosition As Long
    Dim fieldValue As String
    Dim isApplied As Boolean
    Dim languageName As String
    Dim fontSizeText As String
    Dim lineNumbersText As String
    Dim numberGapText As String
    Dim lineHighlightsText As String
    Dim defaultColor As Long
    Dim sourceText As String
    Dim slideText As String
    Dim runRecords As New Collection
    Dim highlightRecords As New Collection
    Dim isNewShape As Boolean

    resultText = Replace(resultText, vbCrLf, vbLf)
    resultText = Replace(resultText, vbCr, vbLf)
    records = Split(resultText, vbLf)

    For recordIndex = LBound(records) To UBound(records)
        recordText = CStr(records(recordIndex))
        separatorPosition = InStr(recordText, "=")

        If separatorPosition > 1 Then
            fieldValue = Mid$(recordText, separatorPosition + 1)

            Select Case Left$(recordText, separatorPosition - 1)
                Case "status"
                    isApplied = fieldValue = "ok"
                Case "language"
                    languageName = UnescapeCodeField(fieldValue)
                Case "fontSize"
                    fontSizeText = fieldValue
                Case "lineNumbers"
                    lineNumbersText = fieldValue
                Case "numberGap"
                    numberGapText = fieldValue
                Case "lineHighlights"
                    lineHighlightsText = fieldValue
                Case "defaultColor"
                    defaultColor = HexToRgb(fieldValue)
                Case "source"
                    sourceText = UnescapeCodeField(fieldValue)
                Case "text"
                    slideText = UnescapeCodeField(fieldValue)
                Case "run"
                    runRecords.Add fieldValue
                Case "highlight"
                    highlightRecords.Add fieldValue
            End Select
        End If
    Next recordIndex

    If Not isApplied Then Exit Sub

    If codeShape Is Nothing Then
        Set codeShape = targetSlide.Shapes.AddTextbox( _
            msoTextOrientationHorizontal, 0, 0, 100, 40)
        codeShape.Name = "PPT Tools code"
        isNewShape = True
    End If

    WriteCodeText codeShape, slideText, Val(fontSizeText), _
        defaultColor, runRecords, highlightRecords

    With codeShape.Tags
        .Add CODE_TAG, "1"
        .Add CODE_SOURCE_TAG, sourceText
        .Add CODE_LANGUAGE_TAG, languageName
        .Add CODE_FONT_SIZE_TAG, fontSizeText
        .Add CODE_LINE_NUMBERS_TAG, lineNumbersText
        .Add CODE_NUMBER_GAP_TAG, numberGapText
        .Add CODE_LINE_HIGHLIGHTS_TAG, lineHighlightsText
    End With

    If isNewShape Then
        With Application.ActivePresentation.PageSetup
            codeShape.Left = IIf(codeShape.Width < .SlideWidth, _
                (.SlideWidth - codeShape.Width) / 2, 0)
            codeShape.Top = IIf(codeShape.Height < .SlideHeight, _
                (.SlideHeight - codeShape.Height) / 2, 0)
        End With
    End If

    codeShape.Select
End Sub

Private Sub WriteCodeText( _
    ByVal codeShape As Shape, _
    ByVal slideText As String, _
    ByVal fontSize As Single, _
    ByVal defaultColor As Long, _
    ByVal runRecords As Collection, _
    ByVal highlightRecords As Collection)

    Dim runRecord As Variant
    Dim runFields As Variant
    Dim highlightRecord As Variant
    Dim highlightFields As Variant

    With codeShape.TextFrame
        .WordWrap = msoFalse
        .AutoSize = ppAutoSizeShapeToFitText

        With .TextRange
            ' Clearing first drops highlights that the new text would inherit.
            .Text = ""
            .Text = slideText
            .ParagraphFormat.Alignment = ppAlignLeft
            .ParagraphFormat.SpaceBefore = 0
            .ParagraphFormat.SpaceAfter = 0

            With .Font
                .Name = CODE_FONT_NAME
                .Size = fontSize
                .Bold = msoFalse
                .Italic = msoFalse
                .Color.RGB = defaultColor
            End With

            ' Each highlight is "start,length,RRGGBB" and covers one padded line.
            ' Highlights go first: setting one gives the whole range the format
            ' of its first character.
            For Each highlightRecord In highlightRecords
                highlightFields = Split(CStr(highlightRecord), ",")
                codeShape.TextFrame2.TextRange.Characters( _
                    CLng(highlightFields(0)), CLng(highlightFields(1))) _
                    .Font.Highlight.RGB = HexToRgb(CStr(highlightFields(2)))
            Next highlightRecord

            ' Each run is "start,length,RRGGBB,bold,italic" with a 1-based start.
            For Each runRecord In runRecords
                runFields = Split(CStr(runRecord), ",")

                With .Characters(CLng(runFields(0)), CLng(runFields(1))).Font
                    .Color.RGB = HexToRgb(CStr(runFields(2)))
                    If runFields(3) = "1" Then .Bold = msoTrue
                    If runFields(4) = "1" Then .Italic = msoTrue
                End With
            Next runRecord
        End With
    End With
End Sub

Private Function HexToRgb(ByVal hexColor As String) As Long
    HexToRgb = RGB( _
        CLng("&H" & Mid$(hexColor, 1, 2)), _
        CLng("&H" & Mid$(hexColor, 3, 2)), _
        CLng("&H" & Mid$(hexColor, 5, 2)))
End Function

Private Function EscapeCodeField(ByVal value As String) As String
    value = Replace(value, "\", "\\")
    value = Replace(value, vbLf, "\n")
    value = Replace(value, vbCr, "\r")
    EscapeCodeField = Replace(value, vbTab, "\t")
End Function

Private Function UnescapeCodeField(ByVal value As String) As String
    Dim scanPosition As Long
    Dim escapePosition As Long
    Dim escapedCharacter As String

    scanPosition = 1

    Do
        escapePosition = InStr(scanPosition, value, "\", vbBinaryCompare)
        If escapePosition = 0 Or escapePosition = Len(value) Then
            UnescapeCodeField = UnescapeCodeField & Mid$(value, scanPosition)
            Exit Function
        End If

        escapedCharacter = Mid$(value, escapePosition + 1, 1)
        Select Case escapedCharacter
            Case "n"
                escapedCharacter = vbLf
            Case "r"
                escapedCharacter = vbCr
            Case "t"
                escapedCharacter = vbTab
        End Select

        UnescapeCodeField = UnescapeCodeField & _
            Mid$(value, scanPosition, escapePosition - scanPosition) & _
            escapedCharacter
        scanPosition = escapePosition + 2
    Loop
End Function
