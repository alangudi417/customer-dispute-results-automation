    ' Helper "Get newest file"
Private Function GetNewestFile(ByVal FolderPath As String) As String

    Dim f As String
    Dim newestFile As String
    Dim newestDate As Date

    f = Dir(FolderPath & "*.xlsm")

    Do While f <> ""

        If FileDateTime(FolderPath & f) > newestDate Then
            newestDate = FileDateTime(FolderPath & f)
            newestFile = FolderPath & f
        End If

        f = Dir

    Loop
    GetNewestFile = newestFile

End Function


'=======================================================================================


    ' Helper "Get today file"
Private Function GetTodayFile(ByVal FolderPath As String) As String

    Dim f As String
    Dim targetName As String
    
    targetName = "UDM_Dispute " & Format(Date, "mm-dd-yyyy") & ".xlsm"

    f = Dir(FolderPath & targetName)

    If f <> "" Then
        GetTodayFile = FolderPath & f
    End If

End Function


'=======================================================================================


    ' Helper: Safe Workbook Finder
Private Function GetWorkbookByKeyword(ByVal keyword As String, ByVal excludeName As String) As Workbook

    Dim wb As Workbook

    For Each wb In Application.Workbooks
        If InStr(1, wb.Name, keyword, vbTextCompare) > 0 Then

            If wb.Name <> excludeName Then
                Set GetWorkbookByKeyword = wb
                Exit Function
            End If

        End If
    Next wb

End Function
