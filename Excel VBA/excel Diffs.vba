Sub highlight()
    Dim xRg1 As Range
    Dim xRg2 As Range
    Dim xTxt As String
    Dim xCell1 As Range
    Dim xCell2 As Range
    Dim I As Long
    Dim J As Integer
    Dim xLen As Integer
    Dim xDiffs As Boolean
    On Error Resume Next
    If ActiveWindow.RangeSelection.Count > 1 Then
      xTxt = ActiveWindow.RangeSelection.AddressLocal
    Else
      xTxt = ActiveSheet.UsedRange.AddressLocal
    End If
lOne:
    Set xRg1 = Application.InputBox("Range A:", "Kutools for Excel", xTxt, , , , , 8)
    If xRg1 Is Nothing Then Exit Sub
    If xRg1.Columns.Count > 1 Or xRg1.Areas.Count > 1 Then
        MsgBox "Multiple ranges or columns have been selected ", vbInformation, "Kutools for Excel"
        GoTo lOne
    End If
lTwo:
    Set xRg2 = Application.InputBox("Range B:", "Kutools for Excel", "", , , , , 8)
    If xRg2 Is Nothing Then Exit Sub
    If xRg2.Columns.Count > 1 Or xRg2.Areas.Count > 1 Then
        MsgBox "Multiple ranges or columns have been selected ", vbInformation, "Kutools for Excel"
        GoTo lTwo
    End If
    If xRg1.CountLarge <> xRg2.CountLarge Then
       MsgBox "Two selected ranges must have the same numbers of cells ", vbInformation, "Kutools for Excel"
       GoTo lTwo
    End If
    xDiffs = (MsgBox("Click Yes to highlight similarities, click No to highlight differences ", vbYesNo + vbQuestion, "Kutools for Excel") = vbNo)
    Application.ScreenUpdating = False
    xRg2.Font.ColorIndex = xlAutomatic
    For I = 1 To xRg1.Count
        Set xCell1 = xRg1.Cells(I)
        Set xCell2 = xRg2.Cells(I)
        If xCell1.Value2 = xCell2.Value2 Then
            If Not xDiffs Then xCell2.Font.Color = vbRed
        Else
            xLen = Len(xCell1.Value2)
            For J = 1 To xLen
                If Not xCell1.Characters(J, 1).Text = xCell2.Characters(J, 1).Text Then Exit For
            Next J
            If Not xDiffs Then
                If J <= Len(xCell2.Value2) And J > 1 Then
                    xCell2.Characters(1, J - 1).Font.Color = vbRed
                End If
            Else
                If J <= Len(xCell2.Value2) Then
                    xCell2.Characters(J, Len(xCell2.Value2) - J + 1).Font.Color = vbRed
                End If
            End If
        End If
    Next
    Application.ScreenUpdating = True
End Sub