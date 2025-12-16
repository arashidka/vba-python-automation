Option Explicit

'============== MASTER CONTROLLER =============='
' Master routine: orchestrates the manuscript build with screen updating preserved
Sub ApplyFullManuscriptTemplate()
    On Error GoTo MasterErr
    Dim doc As Document
    Dim screenUpdateState As Boolean

    screenUpdateState = Application.ScreenUpdating
    Application.ScreenUpdating = False

    Set doc = ActiveDocument

    InitializeDocument doc
    ConfigureStyles doc
    BuildFrontMatter doc
    ApplyBodyTemplate doc
    MapSectionsAndNormalizeHeadings doc
    InsertHeadersFooters doc
    FinalizeDocument doc

CleanExit:
    Application.ScreenUpdating = screenUpdateState
    Exit Sub
MasterErr:
    HandleError "ApplyFullManuscriptTemplate", Err
    Resume CleanExit
End Sub

'============== MODULE: MAP & NORMALIZE HEADINGS =============='
' Maps sections/headings and enforces consistent heading styles
Sub MapSectionsAndNormalizeHeadings(ByVal doc As Document)
    On Error GoTo ErrorHandler

    Dim para As Paragraph
    Dim headingReport As String
    Dim normalized As String
    Dim sectionIndex As Long

    headingReport = "Heading map:" & vbCrLf

    For Each para In doc.Paragraphs
        normalized = NormalizeHeadingParagraph(para)
        If Len(normalized) > 0 Then
            sectionIndex = para.Range.Sections(1).Index
            headingReport = headingReport & "Section " & sectionIndex & ": " & normalized & " (" & para.Range.Style & ")" & vbCrLf
        End If
    Next para

    Debug.Print headingReport
    Exit Sub
ErrorHandler:
    HandleError "MapSectionsAndNormalizeHeadings", Err
End Sub

' Legacy entry point preserved for compatibility
Sub ApplyMurtidaTemplate()
    On Error GoTo LegacyErr
    ApplyFullManuscriptTemplate
    Exit Sub
LegacyErr:
    HandleError "ApplyMurtidaTemplate", Err
End Sub

'============== SUPPORT: ERROR HANDLING =============='
' Centralized error handler to log and display errors
Private Sub HandleError(ByVal procedureName As String, ByVal errObj As ErrObject)
    Dim msg As String
    msg = "Error in " & procedureName & ": " & errObj.Number & " - " & errObj.Description
    Debug.Print msg
    MsgBox msg, vbExclamation, "Murtida Template"
End Sub

'============== MODULE: CLEANUP ORPHANED NUMBERS =============='
' Removes paragraphs that contain only orphaned numbers while preserving TOC and headers/footers
Sub RemoveOrphanedNumbers()
    On Error GoTo ErrorHandler

    Dim doc As Document
    Dim para As Paragraph
    Dim trimmedText As String
    Dim cleaned As String
    Dim removedCount As Long
    Dim undoRecord As UndoRecord

    Set doc = ActiveDocument
    Set undoRecord = Application.UndoRecord

    undoRecord.StartCustomRecord "Remove Orphaned Numbers"

    For Each para In doc.Paragraphs
        If ShouldInspectParagraph(para) Then
            trimmedText = Trim$(para.Range.Text)

            If Len(trimmedText) > 0 And Len(trimmedText) <= 4 Then
                cleaned = NormalizeNumericCandidate(trimmedText)

                If IsNumeric(cleaned) Then
                    para.Range.Delete
                    removedCount = removedCount + 1
                End If
            End If
        End If
    Next para

    MsgBox "Removed " & removedCount & " orphaned number lines", vbInformation, "Remove Orphaned Numbers"

CleanExit:
    undoRecord.EndCustomRecord
    Exit Sub
ErrorHandler:
    HandleError "RemoveOrphanedNumbers", Err
    GoTo CleanExit
End Sub

Private Function ShouldInspectParagraph(ByVal para As Paragraph) As Boolean
    On Error Resume Next

    Dim styleName As String

    ShouldInspectParagraph = False

    If para.Range.StoryType <> wdMainTextStory Then Exit Function
    If para.Range.Information(wdInFootnoteEndnotePane) Then Exit Function
    If para.Range.Information(wdWithInTable) Then Exit Function

    styleName = ParaStyleName(para)
    If styleName Like "TOC*" Or styleName Like "Index*" Then Exit Function

    ShouldInspectParagraph = True
End Function

Private Function ParaStyleName(ByVal para As Paragraph) As String
    On Error Resume Next

    If TypeName(para.Range.Style) = "Style" Then
        ParaStyleName = para.Range.Style.NameLocal
    Else
        ParaStyleName = CStr(para.Range.Style)
    End If
End Function

Private Function NormalizeNumericCandidate(ByVal value As String) As String
    Dim candidate As String

    candidate = Trim$(value)
    candidate = Replace(candidate, vbTab, "")
    candidate = Replace(candidate, Chr(160), "")
    candidate = Replace(candidate, ".", "")
    candidate = Replace(candidate, ":", "")
    candidate = Replace(candidate, ")", "")
    candidate = Replace(candidate, "(", "")
    candidate = Replace(candidate, ";", "")

    NormalizeNumericCandidate = candidate
End Function

'============== MODULE: INITIALIZE DOCUMENT =============='
' Sets global formatting: margins, default section settings, and document properties
Sub InitializeDocument(ByVal doc As Document)
    On Error GoTo ErrorHandler

    With doc.PageSetup
        .TopMargin = InchesToPoints(1)
        .BottomMargin = InchesToPoints(1)
        .LeftMargin = InchesToPoints(1)
        .RightMargin = InchesToPoints(1)
        .DifferentFirstPageHeaderFooter = True
    End With

    doc.BuiltInDocumentProperties(wdPropertyTitle).Value = "Murtida iyo Maadda"
    doc.BuiltInDocumentProperties(wdPropertySubject).Value = "Manuscript Template Automation"

    Exit Sub
ErrorHandler:
    HandleError "InitializeDocument", Err
End Sub

'============== MODULE: CONFIGURE STYLES =============='
' Creates or updates Normal, Heading, and custom styles
Sub ConfigureStyles(ByVal doc As Document)
    On Error GoTo ErrorHandler

    With doc.Styles(wdStyleNormal).Font
        .Name = "Times New Roman"
        .Size = 12
    End With
    With doc.Styles(wdStyleNormal).ParagraphFormat
        .LineSpacingRule = wdLineSpace1pt5
        .SpaceAfter = 12
        .FirstLineIndent = InchesToPoints(0.3)
    End With

    With doc.Styles(wdStyleHeading1)
        .Font.Name = "Times New Roman"
        .Font.Size = 16
        .Font.Bold = True
        .ParagraphFormat.SpaceBefore = 12
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With

    With doc.Styles(wdStyleHeading2)
        .Font.Name = "Times New Roman"
        .Font.Size = 14
        .Font.Bold = True
        .ParagraphFormat.SpaceBefore = 6
        .ParagraphFormat.SpaceAfter = 6
        .ParagraphFormat.Alignment = wdAlignParagraphLeft
    End With

    Dim storyTitleStyle As Style
    Dim lessonStyle As Style

    Set storyTitleStyle = EnsureStyleExists(doc, "Story Title", wdStyleTypeParagraph)
    ConfigureStoryTitleStyle storyTitleStyle

    Set lessonStyle = EnsureStyleExists(doc, "Lesson", wdStyleTypeParagraph)
    ConfigureLessonStyle lessonStyle

    Exit Sub
ErrorHandler:
    HandleError "ConfigureStyles", Err
End Sub

'============== MODULE: BUILD FRONT MATTER =============='
' Builds the Title Page, Dedication/Acknowledgments, TOC, and Preface
Sub BuildFrontMatter(ByVal doc As Document)
    On Error GoTo ErrorHandler

    BuildTitlePage doc
    InsertDedicationSection doc
    GenerateTOCPlaceholder doc
    FormatPrefaceSection doc

    Exit Sub
ErrorHandler:
    HandleError "BuildFrontMatter", Err
End Sub

'============== MODULE: APPLY BODY TEMPLATE =============='
' Adds stories, glossary, and end matter sections
Sub ApplyBodyTemplate(ByVal doc As Document)
    On Error GoTo ErrorHandler

    ApplyStoryTemplate doc
    InsertGlossarySection doc
    BuildEndMatter doc

    Exit Sub
ErrorHandler:
    HandleError "ApplyBodyTemplate", Err
End Sub

'============== MODULE: INSERT HEADERS/FOOTERS =============='
' Adds centered Arabic page numbers to each section while skipping the title page
Sub InsertHeadersFooters(ByVal doc As Document)
    On Error GoTo ErrorHandler

    AddCenteredPageNumbers doc, True

    Exit Sub
ErrorHandler:
    HandleError "InsertHeadersFooters", Err
End Sub

'============== MODULE: FINALIZE DOCUMENT =============='
' Performs light cleanup and notifies the user
Sub FinalizeDocument(ByVal doc As Document)
    On Error GoTo ErrorHandler

    doc.Content.Characters.Last.Select
    Selection.Collapse wdCollapseEnd
    Selection.TypeParagraph
    Selection.HomeKey wdStory

    MsgBox "Murtida template applied successfully.", vbInformation, "Murtida Template"
    Exit Sub
ErrorHandler:
    HandleError "FinalizeDocument", Err
End Sub

'============== SUPPORT: STYLES AND CONTENT BUILDERS =============='
' Adds centered page numbers to each section footer
Private Sub AddCenteredPageNumbers(ByVal doc As Document, Optional ByVal skipFirstPage As Boolean = False)
    On Error GoTo PageErr
    Dim sec As Section
    Dim footer As HeaderFooter
    Dim addOnFirstPage As Boolean

    addOnFirstPage = Not skipFirstPage

    For Each sec In doc.Sections
        sec.PageSetup.DifferentFirstPageHeaderFooter = skipFirstPage
        Set footer = sec.Footers(wdHeaderFooterPrimary)

        If Not footer Is Nothing Then
            ClearPageNumbers footer
            footer.PageNumbers.NumberStyle = wdPageNumberStyleArabic
            footer.PageNumbers.RestartNumberingAtSection = False
            footer.PageNumbers.Add PageNumberAlignment:=wdAlignPageNumberCenter, FirstPage:=addOnFirstPage
        End If
    Next sec
    Exit Sub
PageErr:
    HandleError "AddCenteredPageNumbers", Err
End Sub

Private Sub ClearPageNumbers(ByVal footer As HeaderFooter)
    On Error GoTo NumberErr

    Do While footer.PageNumbers.Count > 0
        footer.PageNumbers(1).Delete
    Loop
    Exit Sub
NumberErr:
    HandleError "ClearPageNumbers", Err
End Sub

' Builds the Title Page with title, author, contact, and date
Sub BuildTitlePage(ByVal doc As Document)
    On Error GoTo TitleErr
    Dim rng As Range
    Set rng = doc.Range(0, 0)
    rng.InsertAfter "Murtida iyo Maadda" & vbCrLf
    rng.Paragraphs(1).Range.Style = wdStyleTitle
    rng.Paragraphs(1).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    rng.InsertAfter vbCrLf & "Author: Mubaarig Farxaan Faarax (Araye)" & vbCrLf
    rng.Paragraphs(2).Range.Font.Size = 14
    rng.Paragraphs(2).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    rng.InsertAfter "Contact: [Email/Phone]" & vbCrLf
    rng.Paragraphs(3).Range.Font.Size = 12
    rng.Paragraphs(3).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter

    rng.InsertAfter "Date: "
    rng.Paragraphs(4).Range.Font.Size = 12
    rng.Paragraphs(4).Range.ParagraphFormat.Alignment = wdAlignParagraphCenter
    rng.Paragraphs(4).Range.Fields.Add rng.Paragraphs(4).Range, wdFieldDate

    rng.InsertAfter vbCrLf & vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    Exit Sub
TitleErr:
    HandleError "BuildTitlePage", Err
End Sub

' Inserts Dedication/Acknowledgments section with placeholder and bulleted list
Sub InsertDedicationSection(ByVal doc As Document)
    On Error GoTo DedicationErr
    Dim rng As Range
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd
    rng.InsertAfter "Dedication" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading1

    rng.InsertAfter "[Your dedication here]" & vbCrLf & vbCrLf
    rng.InsertAfter "Acknowledgments:" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading2

    rng.InsertAfter ChrW(8226) & " Name 1" & vbCrLf & ChrW(8226) & " Name 2" & vbCrLf & ChrW(8226) & " Name 3" & vbCrLf & vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    Exit Sub
DedicationErr:
    HandleError "InsertDedicationSection", Err
End Sub

' Inserts Table of Contents placeholder with headings and page number placeholders
Sub GenerateTOCPlaceholder(ByVal doc As Document)
    On Error GoTo TOCErr
    Dim rng As Range
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd
    rng.InsertAfter "Table of Contents" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading1

    Dim i As Integer
    For i = 1 To 5
        rng.InsertAfter i & ". Story Title " & i & " ............................................. Page X" & vbCrLf
    Next i
    rng.InsertAfter vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    Exit Sub
TOCErr:
    HandleError "GenerateTOCPlaceholder", Err
End Sub

' Inserts Preface section with heading and placeholder text
Sub FormatPrefaceSection(ByVal doc As Document)
    On Error GoTo PrefaceErr
    Dim rng As Range
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd
    rng.InsertAfter "Preface" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading1

    rng.InsertAfter "[Brief introduction to the book’s purpose, themes, and intended audience.]" & vbCrLf & vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    Exit Sub
PrefaceErr:
    HandleError "FormatPrefaceSection", Err
End Sub

' Inserts a sample story template with all formatting and styles
Sub ApplyStoryTemplate(ByVal doc As Document)
    On Error GoTo StoryErr
    Dim rng As Range
    Set rng = doc.Content

    Dim storyTitleStyle As Style
    Dim lessonStyle As Style
    Set storyTitleStyle = EnsureStyleExists(doc, "Story Title", wdStyleTypeParagraph)
    ConfigureStoryTitleStyle storyTitleStyle

    Set lessonStyle = EnsureStyleExists(doc, "Lesson", wdStyleTypeParagraph)
    ConfigureLessonStyle lessonStyle

    rng.Collapse wdCollapseEnd

    ' Thematic section header
    rng.InsertAfter "---" & vbCrLf & "Wisdom Tales" & vbCrLf & "---" & vbCrLf & vbCrLf

    ' Story 1
    rng.InsertAfter "1. HALKAN AYAA KA IFTIIN BADAN" & vbCrLf
    rng.Paragraphs.Last.Range.Style = storyTitleStyle

    rng.InsertAfter "*Sometimes we look for answers in the easiest place, not the right one.*" & vbCrLf & vbCrLf

    rng.InsertAfter "There was once a man named Juxo who lost his ring inside his house. After searching for a long time without success, he went outside and began looking for it under a streetlamp. His neighbors saw him and asked," & vbCrLf
    rng.InsertAfter """Why are you searching here if you lost your ring inside?""" & vbCrLf
    rng.InsertAfter "Juxo replied," & vbCrLf
    rng.InsertAfter """Because there is more light here.""" & vbCrLf & vbCrLf

    rng.InsertAfter "Lesson: It is tempting to look for solutions where it is easiest, but true answers are often found where the problem began." & vbCrLf
    rng.Paragraphs.Last.Range.Style = lessonStyle

    rng.InsertAfter vbCrLf

    ' Story 2
    rng.InsertAfter "2. HOOYO MAXAA AY TIMAHAAGU LA CIRRAYSTEEN" & vbCrLf
    rng.Paragraphs.Last.Range.Style = storyTitleStyle

    rng.InsertAfter "*Children’s actions can affect their parents more than they realize.*" & vbCrLf & vbCrLf

    rng.InsertAfter "A clever young boy noticed that half of his mother’s hair was turning gray. He wondered why, but kept the question to himself." & vbCrLf
    rng.InsertAfter "One day, during a school event, students were allowed to ask their parents questions. The boy finally asked," & vbCrLf
    rng.InsertAfter """Mother, why is your hair turning gray?""" & vbCrLf
    rng.InsertAfter "She replied," & vbCrLf
    rng.InsertAfter """It’s because of your mischief and noise!""" & vbCrLf
    rng.InsertAfter "The boy, quick-witted, responded," & vbCrLf
    rng.InsertAfter """Now I understand why Grandma’s hair is completely gray—it must be because of you!""" & vbCrLf & vbCrLf

    rng.InsertAfter "Lesson: Our actions have consequences for those who care for us, sometimes in ways we don’t realize." & vbCrLf
    rng.Paragraphs.Last.Range.Style = lessonStyle

    rng.InsertAfter vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    Exit Sub
StoryErr:
    HandleError "ApplyStoryTemplate", Err
End Sub

' Ensures a custom style exists and returns it
Private Function EnsureStyleExists(ByVal doc As Document, ByVal styleName As String, ByVal styleType As WdStyleType) As Style
    On Error GoTo CreateStyle
    If StyleExists(doc, styleName) Then
        Set EnsureStyleExists = doc.Styles(styleName)
        Exit Function
    End If
CreateStyle:
    On Error GoTo StyleErr
    Set EnsureStyleExists = doc.Styles.Add(Name:=styleName, Type:=styleType)
    Exit Function
StyleErr:
    HandleError "EnsureStyleExists" & " (" & styleName & ")", Err
End Function

' Checks whether a style exists without raising an error
Private Function StyleExists(ByVal doc As Document, ByVal styleName As String) As Boolean
    On Error Resume Next
    Dim tmp As Style
    Set tmp = doc.Styles(styleName)
    StyleExists = Not tmp Is Nothing
    Set tmp = Nothing
    On Error GoTo 0
End Function

Private Sub ConfigureStoryTitleStyle(ByVal storyTitleStyle As Style)
    If storyTitleStyle Is Nothing Then Exit Sub

    With storyTitleStyle.Font
        .Name = "Times New Roman"
        .Size = 16
        .Bold = True
    End With
    With storyTitleStyle.ParagraphFormat
        .SpaceAfter = 6
        .Alignment = wdAlignParagraphLeft
    End With
End Sub

Private Sub ConfigureLessonStyle(ByVal lessonStyle As Style)
    If lessonStyle Is Nothing Then Exit Sub

    With lessonStyle.Font
        .Name = "Times New Roman"
        .Italic = True
        .Size = 12
    End With
    With lessonStyle.ParagraphFormat
        .SpaceBefore = 6
        .SpaceAfter = 12
        .Alignment = wdAlignParagraphLeft
    End With
End Sub

Private Function NormalizeHeadingParagraph(ByVal para As Paragraph) As String
    On Error GoTo ErrorHandler

    Dim textValue As String
    Dim cleaned As String

    textValue = para.Range.Text
    cleaned = CleanParagraphText(textValue)

    Select Case UCase$(cleaned)
        Case "MURTIDA IYO MAADDA"
            para.Range.Style = wdStyleTitle
            NormalizeHeadingParagraph = "Murtida iyo Maadda"
        Case "DEDICATION"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "Dedication"
        Case "ACKNOWLEDGMENTS", "ACKNOWLEDGMENTS:"
            para.Range.Style = wdStyleHeading2
            NormalizeHeadingParagraph = "Acknowledgments"
        Case "TABLE OF CONTENTS"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "Table of Contents"
        Case "PREFACE"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "Preface"
        Case "WISDOM TALES"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "Wisdom Tales"
        Case "GLOSSARY"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "Glossary"
        Case "ABOUT THE AUTHOR"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "About the Author"
        Case "COPYRIGHT NOTICE"
            para.Range.Style = wdStyleHeading1
            NormalizeHeadingParagraph = "Copyright Notice"
        Case Else
            NormalizeHeadingParagraph = vbNullString
    End Select

    Exit Function
ErrorHandler:
    HandleError "NormalizeHeadingParagraph", Err
End Function

Private Function CleanParagraphText(ByVal value As String) As String
    Dim cleaned As String

    cleaned = Replace(value, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    cleaned = Replace(cleaned, vbTab, " ")
    cleaned = Trim$(cleaned)

    If Right$(cleaned, 1) = ":" Then
        cleaned = Left$(cleaned, Len(cleaned) - 1)
    End If

    CleanParagraphText = cleaned
End Function

' Inserts Glossary section with placeholder terms
Sub InsertGlossarySection(ByVal doc As Document)
    On Error GoTo GlossaryErr
    Dim rng As Range
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd
    rng.InsertAfter "Glossary" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading1

    rng.InsertAfter "- Murti: Wisdom, proverb" & vbCrLf
    rng.InsertAfter "- Maad: Humor, joke" & vbCrLf
    rng.InsertAfter "- Juxo: A common character in Somali folklore, often representing the “everyman” or a trickster" & vbCrLf & vbCrLf
    rng.Collapse wdCollapseEnd
    rng.InsertBreak Type:=wdPageBreak
    Exit Sub
GlossaryErr:
    HandleError "InsertGlossarySection", Err
End Sub

' Inserts About the Author and Copyright Notice sections
Sub BuildEndMatter(ByVal doc As Document)
    On Error GoTo EndMatterErr
    Dim rng As Range
    Set rng = doc.Content
    rng.Collapse wdCollapseEnd

    rng.InsertAfter "About the Author" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading1
    rng.InsertAfter "Mubaarig Farxaan Faarax (Araye) is a Somali writer dedicated to preserving and sharing the wisdom and humor of Somali oral tradition. He welcomes feedback and suggestions at [email address]." & vbCrLf & vbCrLf

    rng.InsertAfter "Copyright Notice" & vbCrLf
    rng.Paragraphs.Last.Range.Style = wdStyleHeading1
    rng.InsertAfter "© 2025 Mubaarig Farxaan Faarax. All rights reserved. No part of this book may be reproduced without permission from the author." & vbCrLf & vbCrLf
    rng.Collapse wdCollapseEnd
    Exit Sub
EndMatterErr:
    HandleError "BuildEndMatter", Err
End Sub
