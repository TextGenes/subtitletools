Attribute VB_Name = "Module1"
Sub RemoveEmptyBlocksInColumn()
    Dim ws As Worksheet
    Dim rng As Range
    Dim rowBlock As Range
    Dim blockRange As Range
    Dim emptyBlockCount As Integer
    Dim startingcolumn As Integer
    Dim columnRange As Range
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet
    ' Set the range from A10 to AI6000
    Set rng = ws.Range("A10:AI6000")
    
    emptyBlockCount = 0
    
    startingcolumn = Application.WorksheetFunction.RoundDown((ActiveCell.Column / 5), 0) * 5 + 1
    Set columnRange = ActiveSheet.Range(Cells(1, startingcolumn), Cells(1, startingcolumn + 4))
         
         
 ' Loop through each row of blocks
For Each rowBlock In rng.Rows
    ' Check if the current row satisfies the conditions
    If rowBlock.Row Mod 5 = 0 Then
        ' Set the current block range
        Set blockRange = rowBlock.Offset(0, startingcolumn - rowBlock.Column).Resize(5, 5)
        
        Debug.Print "Active Range:" & blockRange.Address

        ' Check if the entire block is empty or contains only zeros
        If IsBlockEmpty(blockRange) Then
            ' Increment the counter for consecutive empty blocks
            emptyBlockCount = emptyBlockCount + 1
            ' Delete the entire block
            Debug.Print "Empty blocks deleted in row " & rowBlock.Row
            blockRange.Delete Shift:=xlUp
        Else
            ' Reset the counter if a non-empty block is found
            emptyBlockCount = 0
        End If

        ' If five consecutive empty blocks are found, exit the loop
        If emptyBlockCount = 5 Then Exit For
    End If
Next rowBlock
End Sub

Function IsBlockEmpty(blockRange As Range) As Boolean
    ' Check if the entire block is empty or contains only zeros
    Dim cell As Range

    ' Loop through each cell in the current block
    For Each cell In blockRange.Cells
        ' Set the current cell range
        Set cell = cell.Resize(1, 1)

        If Not IsError(cell.Value) Then
            If Not IsEmpty(cell.Value) Then
                If IsNumeric(cell.Value) Then
                    If cell.Value <> 0 Then
                        ' The cell has a non-zero numeric value
                        IsBlockEmpty = False
                        Exit Function
                    End If
                Else
                    ' The cell has a non-numeric value
                    IsBlockEmpty = False
                    Exit Function
                End If
            End If
        End If
    Next cell

    ' If no non-empty cells are found, the block is considered empty
    IsBlockEmpty = True
End Function
