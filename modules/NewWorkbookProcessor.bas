    ' Process new file
Private Sub ProcessNewWorkbook( _
    ByVal wbNew As Workbook, _
    ByVal wbOld As Workbook, _
    ByVal wbGenerator As Workbook)

    Dim wsNew As Worksheet
    Dim wsOld As Worksheet
    Dim wsMain As Worksheet

    Dim tblNew As ListObject
    Dim tblOld As ListObject

    Dim dictOld As Object

    Dim i As Long
    Dim lastRow As Long

    Dim colCaseID As Long
    Dim colAOR As Long
    Dim colDays As Long
    Dim colAmt As Long

    Dim caseID As String
    Dim AOR As String
    Dim DaysVal As Double
    Dim AmtVal As Double

    Set wsMain = wbGenerator.Worksheets("Main")

    Set wsNew = wbNew.Worksheets("Open")
    Set wsOld = wbOld.Worksheets("Open")

    Set tblNew = wsNew.ListObjects("Open")
    Set tblOld = wsOld.ListObjects("Open")

    Set dictOld = CreateObject("Scripting.Dictionary")


        ' Build Dispute "Case ID Column" Dictionary for Old Board
    colCaseID = GetColumnIndex(tblOld, "Case ID")

    For i = 1 To tblOld.ListRows.Count
        caseID = Trim(CStr(tblOld.DataBodyRange(i, colCaseID).value))

        If Len(caseID) > 0 Then

            If Not dictOld.Exists(caseID) Then
                dictOld.Add caseID, True
            End If

        End If
    Next i


        ' Add headers to columns AE:AG
    wsNew.Columns("AE:AG").ClearContents
    wsNew.Range("AE1").value = "New"
    wsNew.Range("AF1").value = "Open"
    wsNew.Range("AG1").value = "AOR2"

    colCaseID = GetColumnIndex(tblNew, "Case ID")
    colAOR = GetColumnIndex(tblNew, "AOR")
    colDays = GetColumnIndex(tblNew, "Days")
    colAmt = GetColumnIndex(tblNew, "Disputed Amount")

    Dim DictCount As Object
    Dim dictAmount As Object

    Dim dictNewCount As Object
    Dim dictNewAmount As Object

    Dim dict90179 As Object
    Dim dict180 As Object

    Set DictCount = CreateObject("Scripting.Dictionary")
    Set dictAmount = CreateObject("Scripting.Dictionary")

    Set dictNewCount = CreateObject("Scripting.Dictionary")
    Set dictNewAmount = CreateObject("Scripting.Dictionary")

    Set dict90179 = CreateObject("Scripting.Dictionary")
    Set dict180 = CreateObject("Scripting.Dictionary")

    For i = 1 To tblNew.ListRows.Count

        caseID = Trim(CStr(tblNew.DataBodyRange(i, colCaseID).value))
        AOR = Trim(CStr(tblNew.DataBodyRange(i, colAOR).value))

        DaysVal = val(tblNew.DataBodyRange(i, colDays).value)
        AmtVal = val(tblNew.DataBodyRange(i, colAmt).value)


        ' Columns AE/AF/AG for Old/New Board

        If dictOld.Exists(caseID) Then
            wsNew.Range("AE" & i + 1).value = "Old"
        Else
            wsNew.Range("AE" & i + 1).value = "New"
        End If

    wsNew.Range("AF" & i + 1).value = "Open"
    wsNew.Range("AG" & i + 1).value = AOR

        ' All Open Totals
        If Not DictCount.Exists(AOR) Then

            DictCount.Add AOR, 0
            dictAmount.Add AOR, 0

        End If

        DictCount(AOR) = DictCount(AOR) + 1
        dictAmount(AOR) = dictAmount(AOR) + AmtVal


    ' Column New Only
        If wsNew.Range("AE" & i + 1).value = "New" Then

            If Not dictNewCount.Exists(AOR) Then
                dictNewCount.Add AOR, 0
                dictNewAmount.Add AOR, 0

            End If

            dictNewCount(AOR) = dictNewCount(AOR) + 1
            dictNewAmount(AOR) = dictNewAmount(AOR) + AmtVal

        End If


        ' Dispute Aging
        If DaysVal >= 90 And DaysVal < 180 Then

           If Not dict90179.Exists(AOR) Then
                dict90179.Add AOR, 0
            End If

            dict90179(AOR) = dict90179(AOR) + 1

        ElseIf DaysVal >= 180 Then

            If Not dict180.Exists(AOR) Then
                dict180.Add AOR, 0
            End If

           dict180(AOR) = dict180(AOR) + 1

        End If
    Next i


    ' Move results from temp pivot results to Main Generator

    Dim r As Long
    Dim lookupAOR As String

    For r = 13 To 23
        lookupAOR = Trim(CStr(wsMain.Cells(r, "B").value))

        If dictAmount.Exists(lookupAOR) Then

            wsMain.Cells(r, "H").value = dictAmount(lookupAOR)
            wsMain.Cells(r, "I").value = DictCount(lookupAOR)

        End If

        If dictNewCount.Exists(lookupAOR) Then
            wsMain.Cells(r, "P").value = dictNewCount(lookupAOR)
            wsMain.Cells(r, "Q").value = dictNewAmount(lookupAOR)

        End If

        If dict90179.Exists(lookupAOR) Then
            wsMain.Cells(r, "J").value = dict90179(lookupAOR)
        End If

        If dict180.Exists(lookupAOR) Then
            wsMain.Cells(r, "K").value = dict180(lookupAOR)
        End If
    Next r

End Sub