Attribute VB_Name = "Module1"
Sub CheckSubtitleNumbers()
    Dim ws As Worksheet
    Dim subtitleColumn As Long
    Dim subtitleCell As Range
    Dim subtitleContent As Variant
    Dim subtitleNumbers As Collection
    Dim missingSubtitles As Collection
    Dim expectedSubtitleNumber As Long
    Dim i As Long


    'This script goes through all subtitle numbers
    '(current column, row 10 and every following 5th row)
    'ignores empty cells and alerts to any numbers missing between the first (1) and the last one.
    '(might miss any missings at the end).
    
    
    ' Set the worksheet to the active sheet
    Set ws = ActiveSheet

    ' Initialize the collection for subtitle numbers
    Set subtitleNumbers = New Collection
    ' Initialize the collection for missing subtitles
    Set missingSubtitles = New Collection

    ' Set the subtitle column to the active column
    subtitleColumn = ActiveCell.Column

    ' Loop through subtitle rows
    For i = 0 To 1200
        ' Set the current subtitle cell
        Set subtitleCell = ws.Cells(10 + 5 * i, subtitleColumn)

        ' Check if the subtitle cell is not empty
        If Not IsEmpty(subtitleCell.Value) Then
            ' Get the content of the subtitle cell
            subtitleContent = subtitleCell.Value

            ' Check if the content is numeric and greater than 0
            If IsNumeric(subtitleContent) And subtitleContent > 0 Then
                ' Check if the content matches the expected subtitle number
                If subtitleContent <> expectedSubtitleNumber + 1 Then
                    ' Add the missing subtitle numbers to the collection
                    For j = expectedSubtitleNumber + 1 To subtitleContent - 1
                        If Not CollectionContains(missingSubtitles, j) Then
                            missingSubtitles.Add j
                        End If
                    Next j
                End If
                ' Update the expected subtitle number
                expectedSubtitleNumber = subtitleContent
            End If

            ' Add the numeric content to the collection
            subtitleNumbers.Add subtitleContent
        End If
    Next i

    ' Display the missing subtitles in a message box
    If missingSubtitles.Count > 0 Then
        Dim errorMessage As String
        errorMessage = "The following subtitles are missing or have incorrect values in column " & Split(Cells(1, subtitleColumn).Address, "$")(1) & ":" & vbCrLf & vbCrLf

        ' Concatenate the missing subtitles
        For Each Number In missingSubtitles
            errorMessage = errorMessage & "Subtitle " & Number & " missing in row " & 10 + 5 * (Number - 1) & vbCrLf
        Next Number

        MsgBox errorMessage, vbExclamation, "Subtitle Check"
    Else
        MsgBox "All subtitles in the current block are correct!", vbInformation, "Subtitle Check"
    End If
End Sub

Function CollectionContains(coll As Collection, key As Variant) As Boolean
    On Error Resume Next
    CollectionContains = Not coll(key) Is Nothing
    On Error GoTo 0
End Function
