    ' Process old file
Private Sub ProcessOldWorkbook( _
        ByVal wbNew As Workbook, _
        ByVal wbOld As Workbook, _
        ByVal wbGenerator As Workbook)

'    Dim wbNew As Workbook

    Dim wsOld As Worksheet
    Dim wsNew As Worksheet
    Dim wsMain As Worksheet

    Dim tblOld As ListObject
    Dim tblNew As ListObject

    Dim dictNew As Object

    Dim dictClosedCount As Object
    Dim dictClosedAmount As Object

    Dim dict90179 As Object
    Dim dict180 As Object

    Dim colCaseID As Long
    Dim colAOR As Long
    Dim colDays As Long
    Dim colAmt As Long

    Dim i As Long
    Dim r As Long

    Dim caseID As String
    Dim AOR As String

    Dim DaysVal As Double
    Dim AmtVal As Double

    Set wsMain = wbGenerator.Worksheets("Main")

    If wbNew Is Nothing Then Exit Sub

    Set wsOld = wbOld.Worksheets("Open")
    Set wsNew = wbNew.Worksheets("Open")
    Set tblOld = wsOld.ListObjects("Open")
    Set tblNew = wsNew.ListObjects("Open")
    Set dictNew = CreateObject("Scripting.Dictionary")


    ' Build Dispute Dictionary for New Board
    colCaseID = GetColumnIndex(tblNew, "Case ID")

    For i = 1 To tblNew.ListRows.Count
        caseID = Trim(CStr(tblNew.DataBodyRange(i, colCaseID).value))

        If Len(caseID) > 0 Then

            If Not dictNew.Exists(caseID) Then
                dictNew.Add caseID, True
            End If
        End If

    Next i

    ' Add info to Columns AE:AG

    wsOld.Columns("AE:AG").ClearContents

    wsOld.Range("AE1").value = "Closed"
    wsOld.Range("AF1").value = "Old"
    wsOld.Range("AG1").value = "AOR2"

    colCaseID = GetColumnIndex(tblOld, "Case ID")
    colAOR = GetColumnIndex(tblOld, "AOR")
    colDays = GetColumnIndex(tblOld, "Days")
    colAmt = GetColumnIndex(tblOld, "Disputed Amount")

    Set dictClosedCount = CreateObject("Scripting.Dictionary")
    Set dictClosedAmount = CreateObject("Scripting.Dictionary")

    Set dict90179 = CreateObject("Scripting.Dictionary")
    Set dict180 = CreateObject("Scripting.Dictionary")


    For i = 1 To tblOld.ListRows.Count

       caseID = Trim(CStr(tblOld.DataBodyRange(i, colCaseID).value))
        AOR = Trim(CStr(tblOld.DataBodyRange(i, colAOR).value))

        DaysVal = val(tblOld.DataBodyRange(i, colDays).value)
        AmtVal = val(tblOld.DataBodyRange(i, colAmt).value)

        If dictNew.Exists(caseID) Then
            wsOld.Range("AE" & i + 1).value = "Open"

        Else

            wsOld.Range("AE" & i + 1).value = "Closed"
            If Not dictClosedCount.Exists(AOR) Then

                dictClosedCount.Add AOR, 0
                dictClosedAmount.Add AOR, 0
            End If

            dictClosedCount(AOR) = _
                dictClosedCount(AOR) + 1

            dictClosedAmount(AOR) = _
                dictClosedAmount(AOR) + AmtVal

            If DaysVal >= 90 And DaysVal < 180 Then
                If Not dict90179.Exists(AOR) Then
                    dict90179.Add AOR, 0
                End If

                dict90179(AOR) = _
                    dict90179(AOR) + 1

            ElseIf DaysVal >= 180 Then
                If Not dict180.Exists(AOR) Then
                    dict180.Add AOR, 0
                End If

                dict180(AOR) = _
                    dict180(AOR) + 1

            End If
        End If

        wsOld.Range("AF" & i + 1).value = "Old"
        wsOld.Range("AG" & i + 1).value = AOR

    Next i


' Paste the results from the temp pivot to Generator Results

    Dim lookupAOR As String
    For r = 13 To 23

        lookupAOR = Trim(CStr(wsMain.Cells(r, "B").value))

        If dictClosedCount.Exists(lookupAOR) Then
            wsMain.Cells(r, "N").value = _
                dictClosedCount(lookupAOR)

            wsMain.Cells(r, "O").value = _
                dictClosedAmount(lookupAOR)

        End If

        If dict180.Exists(lookupAOR) Then

            wsMain.Cells(r, "R").value = _
                dict180(lookupAOR)

        End If

        If dict90179.Exists(lookupAOR) Then

            wsMain.Cells(r, "S").value = _
                dict90179(lookupAOR)
        End If
    Next r

End Sub