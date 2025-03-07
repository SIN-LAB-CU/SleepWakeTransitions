Sub CalculateSleepWakeTransitions()
    Dim ws As Worksheet, resultSheet As Worksheet
    Dim startRow As Long, thresholdRow As Long
    Dim hourSegmentSize As Long
    Dim cageStartCol As Long, cageEndCol As Long
    Dim i As Long, j As Long, col As Long
    Dim transitions As Long
    Dim thresholdValue As Double
    Dim currentValue As Double, previousValue As Double

    ' Set worksheet and parameters
    Set ws = ThisWorkbook.Sheets("m trans DoD WT males G2 baselin")
    thresholdRow = 2 ' Row containing the threshold values
    startRow = 3     ' Data starts from this row
    hourSegmentSize = 1800 ' Number of rows for each hour segment
    cageStartCol = 2  ' First cage column (e.g., column B)
    cageEndCol = 17   ' Last cage column (e.g., column Q)

    ' Create a new sheet for results
    On Error Resume Next
    Application.DisplayAlerts = False
    ThisWorkbook.Sheets("TransitionsResults").Delete ' Delete old results sheet if it exists
    Application.DisplayAlerts = True
    On Error GoTo 0
    Set resultSheet = ThisWorkbook.Sheets.Add
    resultSheet.Name = "TransitionsResults"

    ' Add headers to the results sheet
    resultSheet.Cells(1, 1).Value = "Hour Segment"
    For col = cageStartCol To cageEndCol
        resultSheet.Cells(1, col - cageStartCol + 2).Value = "Cage " & (col - cageStartCol + 1)
    Next col

    ' Loop through rows in hour segments
    Dim segmentCounter As Long
    segmentCounter = 1 ' Counter for hour segments

    For i = startRow To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row Step hourSegmentSize
        resultSheet.Cells(segmentCounter + 1, 1).Value = "Segment " & segmentCounter

        For col = cageStartCol To cageEndCol
            ' Set the threshold for the current column
            thresholdValue = ws.Cells(thresholdRow, col).Value

            ' Reset transition counter for each column
            transitions = 0

            ' Loop through rows in the current segment
            For j = i To Application.Min(i + hourSegmentSize - 1, ws.Cells(ws.Rows.Count, 1).End(xlUp).Row)
                currentValue = ws.Cells(j, col).Value
                previousValue = ws.Cells(j - 1, col).Value

                ' Count transitions (crossing the threshold in either direction)
                If Not IsEmpty(currentValue) And Not IsEmpty(previousValue) Then
                    If (currentValue > thresholdValue And previousValue < thresholdValue) Or _
                       (currentValue < thresholdValue And previousValue > thresholdValue) Then
                        transitions = transitions + 1
                    End If
                End If
            Next j

            ' Write the result to the results sheet
            resultSheet.Cells(segmentCounter + 1, col - cageStartCol + 2).Value = transitions
        Next col

        segmentCounter = segmentCounter + 1
    Next i

    MsgBox "Transitions calculated and labeled successfully!"
End Sub



