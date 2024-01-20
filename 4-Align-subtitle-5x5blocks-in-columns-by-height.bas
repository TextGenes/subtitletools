Attribute VB_Name = "Module1"



'4 aligning multiple columns (in blocks) vertically
'VBA script
'for excel

'Preparation:
'Make sure that subtitles are formatted regularly, at 5 row intervals, with steps 1 and 3 completed (2 for new speaker timings). 
'Each cell block comprises 5x5 cells, so there should be 5 columns between one timing in seconds and the next one. 
'(columns suggestion: 1 timing in seconds, 2 timing speaker (+names)/subtitle column, 3-5 empty/comments).
'(The first cell with timing in seconds ("0") would be in A10.)
'(Place the first (additional) timing in B11, e.g., "00:00:01" in B11, gives "1" in A11)

'This script runs for cells A10 to AI6000. 
'Up to ten subtitle columns (or 5 languages with two versions) can be posted. For more, change to higher letters for: AI6000. If it outruns row 6000, change to more rows.

'Function:
'Comparisons are performed on the respective cell (2nd row, 1st column) in each block (subtitle start timing in seconds), if it contains a number higher than 0 (all others ignored).
'In each block, the scripts picks comparison cell (timing in seconds) and compares it to others in the row (if there is more than one timing in this row).
'If any timings are more than 2 seconds higher than others in the row, the higher timings are moved down a block, with comments and all, by inserting a 5x5 empty cell block.
'At the end, you may pull along the numbers formula from step 3 (but not needed, all empty are 0).

'For repeated use:
'You can copy in subtitles in another language or version in the next 5x5 block.
'Else, when copying a subtitle column back into notepad and saving as .srt, when you upload to youtube, the empty rows will be cleared.
'(So, if you redownload an .srt, repeat steps 1 (oneliners in program), 3(pull up formula in seconds) and run 4 (height alignment) to realign).




Sub SortAndShiftBlocks()
    Dim ws As Worksheet
    Dim rng As Range
    Dim rowBlock As Range
    Dim blockRange As Range
    Dim compareValue As Double
    Dim sensitivity As Double
    Dim referencevalue As Double
    Dim candidateCollection As Collection
    Dim comparisonCollection As Collection
    Dim targetCell As Range
    Dim emptyBlock As Range
    Dim i As Long, j As Long, k As Long, l As Long, m As Long

    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    'alternative for precise: Set ws = ThisWorkbook.Sheets ("EP1-prefinal") ' Change to your actual sheet name

    ' Set the range from A10 to AI6000
    Set rng = ws.Range("A10:AI6000")

    ' Set sensitivity
    sensitivity = 2

    ' Loop through each row of blocks
    For Each rowBlock In rng.Rows
    
     ' Check if the current row satisfies the conditions
        If rowBlock.Row Mod 5 <> 0 Then
            ' Your processing code here
            GoTo NextRowBlock
        End If
        

        ' Print the current row number for debugging purposes
        Debug.Print "Current Row: " & rowBlock.Row

        ' Initialize candidate collection
        Set candidateCollection = New Collection

        ' Loop through each 5x5 block within the row
        For j = 1 To rowBlock.Columns.Count Step 5 ' Start from the 1st column, increment by 5
            ' Set the current block range
            Set blockRange = rowBlock.Columns(j).Resize(5, 5)
            

            ' Compare each comparison value to find those above 0
            For i = 1 To blockRange.Columns.Count Step 5
                If Not IsEmpty(blockRange.Cells(2, i).Value) Then
                    If IsNumeric(blockRange.Cells(2, i).Value) And blockRange.Cells(2, i).Value > 0 Then
                        compareValue = blockRange.Cells(2, i).Value
                        candidateCollection.Add compareValue
                    End If
                End If
            Next i
        Next j

        ' If there are fewer than 2 items in the candidate collection, or if the comparison collection is empty, move to the next row block
        If candidateCollection.Count < 2 Then
            Set candidateCollection = Nothing
            Set comparisonCollection = Nothing
            referencevalue = 0
            GoTo NextRowBlock
        End If

        ' Find the lowest value in the candidate collection
        Dim minValue As Double
        minValue = candidateCollection(1)

        ' Loop through the candidate collection to find the minimum value, that is higher than 0
        For k = 2 To candidateCollection.Count
            If 0 < candidateCollection(k) And candidateCollection(k) < minValue Then
                minValue = candidateCollection(k)
            End If
        Next k

        ' Set the reference value as the minimum value plus sensitivity
        referencevalue = minValue + sensitivity

        ' Initialize comparison collection
        Set comparisonCollection = New Collection

        ' Loop through each 5x5 block within the row again
        For l = 1 To rowBlock.Columns.Count Step 5 ' Start from the 1st column, increment by 5
            ' Set the current block range
            Set blockRange = rowBlock.Columns(l).Resize(5, 5)

            ' Compare each comparison value to the reference value
            For m = 1 To blockRange.Columns.Count Step 5
                If Not IsEmpty(blockRange.Cells(2, m).Value) Then
                    If IsNumeric(blockRange.Cells(2, m).Value) And blockRange.Cells(2, m).Value > 0 Then
                        compareValue = blockRange.Cells(2, m).Value

                        ' Check if comparison value is higher than reference value
                        If compareValue > referencevalue Then
                            ' Add the cell to the collection
                            Set targetCell = blockRange.Cells(2, m)
                            comparisonCollection.Add targetCell
                        End If
                    End If
                End If
            Next m
        Next l

        ' If the comparison collection is empty, move to the next row block
        If comparisonCollection.Count = 0 Then
            Set candidateCollection = Nothing
            Set comparisonCollection = Nothing
            referencevalue = 0
            GoTo NextRowBlock
        End If

       ' Set the active cell to the first column of the current row
       rowBlock.Cells(1, 1).Activate
       
       ' If there are cells above the reference value, move and insert the empty block
       ' Move to column one and process the collection
       For Each targetCell In comparisonCollection
          'Go to the target cell
          targetCell.Activate
          ' Shift the 5x5 block starting from targetCell down
          targetCell.Resize(5, 5).Offset(-1, 0).Insert Shift:=xlDown
       Next targetCell


NextRowBlock:
        ' Clear the comparison collection for the next row
        Set comparisonCollection = Nothing

        ' Clear the candidate collection for the next row
        Set candidateCollection = Nothing
        
        'reset referencevalues and counters to be sure
        referencevalue = 0
        minValue = 0
        i = 0
        j = 0
        k = 0
        l = 0
        m = 0
        
        ' Set the active cell to the first column of the current row
       rowBlock.Cells(1, 1).Activate

        'Move down one more row to restart in next row
    Next rowBlock
End Sub
