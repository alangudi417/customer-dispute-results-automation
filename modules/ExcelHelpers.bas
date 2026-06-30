    ' Helper "Get the table from Column Index
Private Function GetColumnIndex( _
    ByVal tbl As ListObject, _
    ByVal HeaderName As String) As Long

    Dim i As Long
    For i = 1 To tbl.ListColumns.Count

        If Trim(tbl.ListColumns(i).Name) = HeaderName Then
            GetColumnIndex = i
            Exit Function
        End If

    Next i

    Err.Raise vbObjectError + 1000, , _
        "Column not found: " & HeaderName

End Function


'============================================================================


    ' Helper: Safe cell write
Private Sub SafeWrite(ByVal ws As Worksheet, ByVal rowNum As Long, ByVal colLetter As String, ByVal dict As Object, ByVal key As String)

    If dict.Exists(key) Then
        ws.Cells(rowNum, colLetter).value = dict(key)
    Else
        ws.Cells(rowNum, colLetter).value = 0
    End If

End Sub


'============================================================================


    ' Delete temp sheet
    ' As the pivot table will go temp during the macro, and deleted before the macro ends
    
Private Sub DeleteSheetIfExists( _
    ByVal wb As Workbook, _
    ByVal SheetName As String)

    On Error Resume Next

    wb.Worksheets(SheetName).Delete

    On Error GoTo 0

End Sub