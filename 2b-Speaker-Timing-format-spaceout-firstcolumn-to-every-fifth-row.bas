Attribute VB_Name = "Module1"
Sub SpaceOutSpeakerTimingsAddFourEmptyCellsBetween()
    
    Dim ws As Worksheet
    Dim lastRow As Long
    Dim currentRow As Long
    Dim currentColumn As String
    
    'Script adds four empty cells between each row of timings in column A, preparation for 5x5 block
    ' Set the worksheet
    Set ws = ActiveSheet
    
    ' Set ws = ThisWorkbook.Sheets("...")
    ' Change "Sheet1" to your actual sheet name
    
    ' Set the column with speaker timings
    currentColumn = "A"
    ' Change "A" to your actual column letter
    
    ' Find the last row with data in the specified column
    lastRow = ws.Cells(ws.Rows.Count, currentColumn).End(xlUp).Row
    
    ' Loop through each row starting from the last row
    For currentRow = lastRow To 2 Step -1
        ' Insert four empty cells below each timing
        ws.Range(currentColumn & currentRow + 1 & ":" & currentColumn & currentRow + 4).Insert Shift:=xlDown
    Next currentRow
End Sub
