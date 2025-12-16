Option Explicit

'============== STANDALONE MAPPING MODULE =============='
' This module provides a standalone section mapping tool that can be executed
' independently to map and normalize headings in a Word document.
'
' Usage: Run MapSectionsStandalone() from VBA editor or assign to a button/shortcut
'
' Purpose: Maps all sections and headings in the active document, normalizes heading
' styles to match standard manuscript formatting, and outputs a report to the
' Immediate Window.

'============== MAIN ENTRY POINT =============='
' Standalone entry point: Maps sections and normalizes headings in the active document
Sub MapSectionsStandalone()
    On Error GoTo ErrorHandler
    
    Dim doc As Document
    Dim para As Paragraph
    Dim headingReport As String
    Dim normalized As String
    Dim sectionIndex As Long
    Dim headingCount As Long
    
    ' Get the active document
    Set doc = ActiveDocument
    
    If doc Is Nothing Then
        MsgBox "No active document found. Please open a document first.", vbExclamation, "Standalone Mapping"
        Exit Sub
    End If
    
    ' Initialize the report
    headingReport = "=== DOCUMENT HEADING MAP ===" & vbCrLf & vbCrLf
    headingCount = 0
    
    ' Iterate through all paragraphs in the document
    For Each para In doc.Paragraphs
        normalized = NormalizeHeadingParagraph(para)
        If Len(normalized) > 0 Then
            headingCount = headingCount + 1
            sectionIndex = para.Range.Sections(1).Index
            headingReport = headingReport & "Section " & sectionIndex & ": " & _
                           normalized & " (" & para.Range.Style & ")" & vbCrLf
        End If
    Next para
    
    ' Add summary to report
    headingReport = headingReport & vbCrLf & "=== SUMMARY ===" & vbCrLf
    headingReport = headingReport & "Total headings found and normalized: " & headingCount & vbCrLf
    headingReport = headingReport & "Total sections: " & doc.Sections.Count & vbCrLf
    
    ' Output report to Immediate Window and show confirmation
    Debug.Print headingReport
    MsgBox "Heading mapping complete. " & headingCount & " heading(s) normalized." & vbCrLf & _
           "Check the Immediate Window (Ctrl+G) for full report.", vbInformation, "Standalone Mapping"
    
    Exit Sub
ErrorHandler:
    HandleError "MapSectionsStandalone", Err
End Sub

'============== HEADING NORMALIZATION =============='
' Normalizes a paragraph if it contains a recognized heading
' Returns the normalized heading text, or empty string if not a recognized heading
Private Function NormalizeHeadingParagraph(ByVal para As Paragraph) As String
    On Error GoTo ErrorHandler
    
    Dim textValue As String
    Dim cleaned As String
    
    textValue = para.Range.Text
    cleaned = CleanParagraphText(textValue)
    
    ' Map recognized headings to their normalized form and appropriate style
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
        Case "TABLE OF CONTENTS", "TOC"
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
            ' Not a recognized heading - return empty string
            NormalizeHeadingParagraph = vbNullString
    End Select
    
    Exit Function
ErrorHandler:
    HandleError "NormalizeHeadingParagraph", Err
    NormalizeHeadingParagraph = vbNullString
End Function

'============== TEXT CLEANING =============='
' Cleans paragraph text by removing line breaks, tabs, and extra whitespace
Private Function CleanParagraphText(ByVal value As String) As String
    Dim cleaned As String
    
    ' Remove line breaks and tabs
    cleaned = Replace(value, vbCr, "")
    cleaned = Replace(cleaned, vbLf, "")
    cleaned = Replace(cleaned, vbTab, " ")
    
    ' Trim whitespace
    cleaned = Trim$(cleaned)
    
    ' Remove trailing colon if present
    If Len(cleaned) > 0 And Right$(cleaned, 1) = ":" Then
        cleaned = Left$(cleaned, Len(cleaned) - 1)
    End If
    
    CleanParagraphText = cleaned
End Function

'============== ERROR HANDLING =============='
' Centralized error handler to log and display errors
Private Sub HandleError(ByVal procedureName As String, ByVal errObj As ErrObject)
    Dim msg As String
    msg = "Error in " & procedureName & ": " & errObj.Number & " - " & errObj.Description
    Debug.Print msg
    MsgBox msg, vbExclamation, "Standalone Mapping Error"
End Sub
