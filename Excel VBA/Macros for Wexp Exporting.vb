' modPostWEXPTools
' ================
' Provides an module with a starting point for the
'  Post WEXP Tools suite.
'
' Copyright Thales Avionics Ltd.
' Written by Elliot Ali, Systems Engineering, Raynes Park.
' Version 1.0 - 14th November 2006

'PostWEXPTools
'This is the actual macro available to Word.
Public Sub PostWEXPTools()
    'Show the WEXPTools dialog
    ' (This will then create a WEXPTools session and do all the
    ' remaining work)
    frmPostExport.Show
End Sub

' modPostWEXPTools
' ================
' Fixes problems caused by the WEXP tool in DOORS.
' These macros will be called by the exporter.
'
' Copyright Thales Avionics Ltd.
' Written by Elliot Ali, Systems Engineering, Raynes Park.
' Version 1.0 - 15th January 2008


'Corrects any problems with the DOORS Section creation process, including
' pages being inserted of the wrong size, headers being modified

'  USE SUPERCEEDED VERSION OF MACRO BELOW, UNLESS YOU HAVE ANY ISSUES
'  A. MOGHUL

Public Sub FixPageSizes_Original()
    Dim s As Section
    Dim a As Variant
    
    Dim t As Table
    Dim v As Table
    Dim r As Row
    Dim w As Row
    Dim c As Cell
    Dim l As String
    
    For Each s In ActiveDocument.Sections
        'Remember orientation, fix size and reapply orientation
        a = s.PageSetup.Orientation
        s.PageSetup.PaperSize = wdPaperA4
        s.PageSetup.Orientation = a
        
        'Fix header tables
        s.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        s.Headers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        s.Headers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidth = 100
        
        'Fix footer tables
        s.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
        's.Footers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        's.Footers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidth = 100
        
        'Stop Word from locking up
        DoEvents
    Next s

    'In tables of 5 columns,
    'Grey out rows with N/A in the fourth cell.
    For Each t In ActiveDocument.Tables
        If t.Columns.Count = 5 Then
            For Each r In t.Rows
                l = r.Cells(4).Range.Paragraphs(1).Range.Text
                If Len(l) = 5 And Left(l, 3) = "N/A" Then
                    r.Cells.Shading.Texture = wdTextureNone
                    r.Cells.Shading.ForegroundPatternColor = wdColorAutomatic
                    r.Cells.Shading.BackgroundPatternColor = wdColorGray10
                    Else
                    r.Cells.Shading.BackgroundPatternColor = wdColorWhite
                    
                End If
            Next r
        End If
        'Stop Word from locking up
        DoEvents
    Next t
 
    
    'In tables of 5 columns,
    'White out rows with text in the third cell.
    For Each v In ActiveDocument.Tables
       If v.Columns.Count = 5 Then
           For Each w In v.Rows
                l = w.Cells(3).Range.Paragraphs(1).Range.Text
                If Len(l) > 5 And Left(l, 8) <> "Expected" Then
                    w.Cells.Shading.Texture = wdTextureNone
                    w.Cells.Shading.ForegroundPatternColor = wdColorAutomatic
                    w.Cells.Shading.BackgroundPatternColor = wdColorWhite
                End If
            Next w
        End If
        'Stop Word from locking up
        DoEvents
    Next v
 
    
End Sub




' modPostWEXPTools
' ================
' Fixes problems caused by the WEXP tool in DOORS.
' These macros will be called by the exporter.
'
' Copyright Thales Avionics Ltd.
' Written by Elliot Ali, Systems Engineering, Raynes Park.
' Version 1.0 - 15th January 2008


'Corrects any problems with the DOORS Section creation process, including
' pages being inserted of the wrong size, headers being modified
'
' A. Moghul 12 June 2013
' Modified to add colour for Pass Fail, also vertically center text in cells.
' If you encounter any problems, use the FixPageSizes_OLD macro

Public Sub FixPageSizes()
    Dim s As Section
    Dim a As Variant
    
    Dim t As Table
    Dim v As Table
    Dim r As Row
    Dim w As Row
    Dim c As Cell
    Dim l As String
    
    For Each s In ActiveDocument.Sections
        'Remember orientation, fix size and reapply orientation
        a = s.PageSetup.Orientation
        s.PageSetup.PaperSize = wdPaperA4
        s.PageSetup.Orientation = a
        
        'Fix header tables
        s.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        s.Headers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        s.Headers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidth = 100
        s.Headers(wdHeaderFooterPrimary).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        
        s.Headers(wdHeaderFooterPrimary).Range.Cells(1).Range.Paragraphs(1).Alignment = wdAlignParagraphCenter
        
        'Fix footer tables
        s.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
        s.Footers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        s.Footers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidth = 100
        
        'Stop Word from locking up
        DoEvents
    Next s

 'Grey out rows with N/A in the fourth cell.
    For Each t In ActiveDocument.Tables
        If t.Columns.Count = 5 Then
            For Each r In t.Rows
                l = r.Cells(4).Range.Paragraphs(1).Range.Text
                
                ' Vertically center all the cells.  There is probably a better way to do this!
                r.Cells(1).Range.Paragraphs(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                r.Cells(2).Range.Paragraphs(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                r.Cells(3).Range.Paragraphs(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                r.Cells(4).Range.Paragraphs(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                r.Cells(5).Range.Paragraphs(1).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
                
                If Len(l) = 5 And Left(l, 3) = "N/A" Then
                    r.Cells.Shading.Texture = wdTextureNone
                    r.Cells.Shading.ForegroundPatternColor = wdColorAutomatic
                    r.Cells.Shading.BackgroundPatternColor = wdColorGray10
                    r.Cells(4).Range.Paragraphs(1).Range.Font.ColorIndex = wdBlack
                Else
                    r.Cells.Shading.BackgroundPatternColor = wdColorWhite
                    r.Cells(4).Range.Paragraphs(1).Range.Font.ColorIndex = wdBlack
                End If
                
                If Len(l) = 6 And Left(l, 4) = "Fail" Then
                    r.Cells(4).Shading.ForegroundPatternColor = wdColorRed
                End If
                
                If Len(l) >= 6 And Left(l, 5) = "Pass/" Then
                ' do nothing as this is the title
                ElseIf Len(l) >= 6 And Left(l, 4) = "Pass" Then
                    r.Cells(4).Shading.ForegroundPatternColor = wdColorBrightGreen
                End If

            Next r
        End If
        'Stop Word from locking up
        DoEvents
    Next t
 
    
    'In tables of 5 columns,
    'White out rows with text in the third cell.
    For Each v In ActiveDocument.Tables
       If v.Columns.Count = 5 Then
           For Each w In v.Rows
                l = w.Cells(3).Range.Paragraphs(1).Range.Text
                If Len(l) > 5 And Left(l, 8) <> "Expected" Then
                    w.Cells.Shading.Texture = wdTextureNone
                    w.Cells.Shading.ForegroundPatternColor = wdColorAutomatic
                    w.Cells.Shading.BackgroundPatternColor = wdColorWhite
                End If
            Next w
        End If
        'Stop Word from locking up
        DoEvents
    Next v
 
     ' MsgBox ("Finished Running FixPageSizes Macro")
    
End Sub




' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'   A.Moghul   January 2014
'
'  Colours page background.  Use this to colour the page white before beginning.
'  Note - there seems to be an issue with this, in that it doesnt colour the page white, unless it is painted
' another colour first manually
'
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Sub AM_SetPageColour(r As Integer, g As Integer, b As Integer)

    ActiveDocument.Background.Fill.Visible = msoTrue
    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(r, g, b)
    ActiveDocument.Background.Fill.Solid
    
End Sub


' Amjad Moghul - Dec 2013
' Sets the Header and footers to be horizontally equally spaced sections

Public Sub AM_2FixSections()

    Dim s As Section
    Dim a As Variant
    
    Dim t As Table
    Dim v As Table
    Dim r As Row
    Dim w As Row
    Dim c As Cell
    Dim l As String
    Dim cnt As Integer
    

    For Each s In ActiveDocument.Sections
        'Remember orientation, fix size and reapply orientation
        a = s.PageSetup.Orientation
        s.PageSetup.PaperSize = wdPaperA4
        s.PageSetup.Orientation = a
        
        'Fix header tables
        s.Headers(wdHeaderFooterPrimary).LinkToPrevious = False
        s.Headers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        s.Headers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidth = 100
        s.Headers(wdHeaderFooterPrimary).Range.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        s.Headers(wdHeaderFooterPrimary).Range.Cells(1).Range.Paragraphs(1).Alignment = wdAlignParagraphCenter

        'Fix footer tables
        s.Footers(wdHeaderFooterPrimary).LinkToPrevious = False
        s.Footers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidthType = wdPreferredWidthPercent
        s.Footers(wdHeaderFooterPrimary).Range.Tables(1).PreferredWidth = 100

        'Stop Word from locking up
        DoEvents
        cnt = cnt + 1
        
    Next s
    
    MsgBox ("Finished. Fixed " & cnt & " Sections")
End Sub





' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'   A.Moghul   January 2014
'
'  Centres all table cells vertically for visual appearance
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Sub AM_3CentreAllTables()

  Dim t As Table
  Dim cnt As Integer
    
   For Each t In ActiveDocument.Tables
        
        t.Select
        Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
        Selection.Cells.VerticalAlignment = wdCellAlignVerticalCenter
        cnt = cnt + 1
    Next t
    
    MsgBox ("Finished centering " & cnt & " Tables")
 
End Sub


' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'   A.Moghul   January 2014
'
'  Dont allow table cells  to break over pages - for visual display only
'
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Sub AM_4TableNoBreak()
    
    Dim t As Table
    Dim cnt As Integer
    

    For Each t In ActiveDocument.Tables
        'If t.Columns.Count = 5 Then
            ' t.Rows.VerticalPosition = wdAlignRowCenter
            t.Rows.AllowBreakAcrossPages = False
            cnt = cnt + 1
            
           ' ActiveDocument.Tables(1).Rows.AllowBreakAcrossPages = False
        'End If
    Next
    
    MsgBox ("Finished. Fixed " & cnt & " Sections")
        
End Sub

' Amjad Moghul - Dec 2013
' Gets rid of extra row from the STR tables .  There is probably a much better way of doing this!

Public Sub AM_5DeleteExtraTableRow()
    Dim s As Section
    Dim a As Variant
    
    Dim t As Table
    Dim v As Table
    Dim r As Row
    Dim w As Row
    Dim c As Cell
    Dim c1 As Integer
    Dim c2 As Integer
    Dim c3 As Integer
    Dim c4 As Integer
    Dim c5 As Integer
    Dim c6 As Integer
    Dim c7 As Integer
    Dim cTot As Integer
    Dim delCnt As Integer
    

    delCnt = 0
    

 'Grey out rows with N/A in the fourth cell.
    For Each t In ActiveDocument.Tables
        If t.Columns.Count = 7 Then
            For Each r In t.Rows
                c1 = Len(r.Cells(1).Range.Paragraphs(1).Range.Text) ' probably a much better way of doing this!
                c2 = Len(r.Cells(2).Range.Paragraphs(1).Range.Text)
                c3 = Len(r.Cells(3).Range.Paragraphs(1).Range.Text)
                c4 = Len(r.Cells(4).Range.Paragraphs(1).Range.Text)
                c5 = Len(r.Cells(5).Range.Paragraphs(1).Range.Text)
                c6 = Len(r.Cells(6).Range.Paragraphs(1).Range.Text)
                c7 = Len(r.Cells(7).Range.Paragraphs(1).Range.Text)
                cTot = c1 + c2 + c3 + c4 + c5 + c6 + c7
               
                If (cTot < 15) Then ' each cell has a space and CR - meaning 2 characters -  7 * 2 + 1 = 15
                     r.Cells.Shading.BackgroundPatternColor = wdColorRed
                    delCnt = delCnt + 1
                      r.Cells.Delete
                End If

            Next r
        End If
        'Stop Word from locking up
        DoEvents
    Next t
    MsgBox ("Finished Running Fix Macro. Fixed " & delCnt & " Table Rows")
    
End Sub




' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'
'   A.Moghul   January 2014
'   FixPageMacros modified.  Rather than colour each row, It is better to colour everything in White, then
'   Change all Passes to Green, All Fails to Red and all N/A (where there is no evidence) to Grey
'
'   If anything doesnt work, you can always resort to the original Macro - FixPageSizes_Original()
'
' /////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Public Sub AM_6NewFixPageSize()
    Dim s As Section
    Dim a As Variant
    
    Dim t As Table
    Dim v As Table
    Dim r As Row
    Dim w As Row
    Dim c As Cell
    Dim l As String
    Dim l3 As String
    Dim i3 As Integer
    
    Dim aCell As Cell
    Dim bEvidence As Boolean

    ' Colour everything White
    AM_SetPageColour 255, 255, 255
    
    ' Centre all Table Cells vertically
   'AM_CentreAllTables
    '
    'Grey out rows with N/A in the fourth cell.
 '
    For Each t In ActiveDocument.Tables
        If t.Columns.Count = 5 Then
            For Each r In t.Rows

                l = r.Cells(4).Range.Paragraphs(1).Range.Text
                l3 = r.Range.Paragraphs.Item(3).Range.Text
                i3 = Len(l3)
                
                ' l = aCell.Range.Paragraphs(1).Range.Text
                             
'                                r.Cells.VerticalAlignment = wdCellAlignVerticalCenter                        ' Align the cells vertically   (Now done in  AM_CentreAllTables )
                 
                bEvidence = False
                If (Left(r.Range.Paragraphs.Item(3).Range.Text, 8) = "Expected") Then         ' Leave the Headers alone
                 DoEvents   ' do nothing - this is the title
                Else
                    If (Left(r.Range.Paragraphs.Item(4).Range.Text, 3) = "N/A") Then          'Turn all N/A Grey
                        r.Cells.Shading.Texture = wdTextureNone
                        If (i3 <= 2) Then                                                     ' The field is not completely empty, so check for len > 2
                                r.Cells.Shading.BackgroundPatternColor = wdColorGray10
                        'r.Cells(4).Range.Paragraphs(1).Range.Font.ColorIndex = wdBlack
                        End If
                          ' Otherwise if it contains expected results text, it is already white
                    End If
 
                    If (Left(r.Range.Paragraphs.Item(4).Range.Text, 4) = "Pass") Then          'Turn all PASS Green
                        r.Cells(4).Range.Paragraphs(1).Range.Font.ColorIndex = wdGreen
                       ' r.Cells(4).Shading.ForegroundPatternColor = wdColorBrightGreen
                        DoEvents
                    ElseIf (Left(r.Range.Paragraphs.Item(4).Range.Text, 4) = "Fail") Then      'Turn all FAIL Red

                        r.Cells(4).Range.Paragraphs(1).Range.Font.ColorIndex = wdRed
                        'r.Cells(4).Shading.ForegroundPatternColor = wdColorRed
                        DoEvents
                    Else
                       'r.Cells.Shading.BackgroundPatternColor = wdColorGray10                  ' Anything not caught, col it grey
                       'r.Cells.Shading.BackgroundPatternColor = wdColorBlue                   ' this was causing incorrect flagging.
                       'r.Cells(4).Range.Paragraphs(1).Range.Font.ColorIndex = wdBlack
                        DoEvents
                    End If
                End If
            Next r
        End If
        'Stop Word from locking up
        DoEvents
    Next t

    MsgBox ("Finished Running FixPageSizes Macro")
    
End Sub




Sub AM_AddLandscape()
'
' AM_AddLandscape Macro
'
'
    ActiveDocument.Range(Start:=Selection.Start, End:=Selection.Start). _
        InsertBreak Type:=wdSectionBreakNextPage
    Selection.Start = Selection.Start + 1
    With ActiveDocument.Range(Start:=Selection.Start, End:=ActiveDocument. _
        Content.End).PageSetup
        .Orientation = wdOrientLandscape
    End With
    Selection.InsertBreak Type:=wdPageBreak
End Sub



Public Sub AM_Temp()
    ' Colour everything White
    ActiveDocument.Background.Fill.Visible = msoTrue
    ActiveDocument.Background.Fill.ForeColor.RGB = RGB(255, 255, 255)
    ActiveDocument.Background.Fill.Solid
End Sub

