Attribute VB_Name = "Module1"
Sub FormatSpeakerTimingsToAddLeadingZero()
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim currentRow As Long

    'Script adds additional leading 0 to timing in format 0:00:00.000
    
    'Set the worksheet
     Set ws = ActiveSheet

    'Set ws = ThisWorkbook.Sheets("...")
    ' Change "Sheet1" to your actual sheet name
    

    ' Find the last row in current column (replace by number of column with the timings; typically only needed for speaker timings, as yt timings have a fixed format)
    lastRow = ws.Cells(ws.Rows.Count, currentCol).End(xlUp).Row
    
    ' Loop through each row in column A
    For currentRow = 1 To lastRow
        ' Check if the time string starts with one digit followed by ":"
        If Left(ws.Cells(currentRow, currentCol).Value, 2) = "0:" Then
            ' Add leading zero if necessary
            ws.Cells(currentRow, currentCol).Value = "0" & ws.Cells(currentRow, currentCol).Value
        End If
    Next currentRow
End Sub
