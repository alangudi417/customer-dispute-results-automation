    ' MAIN MACRO
Sub dispute_results()

    'Dim workbooks
    Dim wbGenerator As Workbook
    Dim wbNew As Workbook
    Dim wbOld As Workbook

    Dim wsMain As Worksheet

    Dim newFolder As String
    Dim oldFolder As String

    Dim newFile As String
    Dim oldFile As String

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual

    On Error GoTo ErrHandler

    Set wbGenerator = ThisWorkbook
    Set wsMain = wbGenerator.Worksheets("Main")

        ' Folder Paths:
    newFolder = "/sample_data/new"
    oldFolder = "/sample_data/old"


        ' Set today's file from New folder
    newFile = GetNewestFile(newFolder)

        ' If, in case is not found
    If newFile = "" Then
        MsgBox "Today's UDM file was not found in New folder.", vbCritical
        GoTo ExitHandler
    End If


        ' Set newest file from Old folder
    oldFile = GetNewestFile(oldFolder)

        ' If, in case is not found
    If oldFile = "" Then
        MsgBox "No file found in Old folder.", vbCritical
        GoTo ExitHandler
    End If


        ' Set workbooks
    Set wbNew = Workbooks.Open(newFile)
    Set wbOld = Workbooks.Open(oldFile)


    ' 1. RESET GENERATOR
    With wsMain
        .Range("D12:G23").value = .Range("E12:H23").value
        .Range("H13:K23").ClearContents
        .Range("N13:S23").ClearContents
        .Range("H12").value = Date
        .Range("H12").NumberFormat = "mm/dd/yyyy"
    End With


    ' 2. PROCESS NEW FILE
    ProcessNewWorkbook wbNew, wbOld, wbGenerator

    wbNew.Save

    ' 3. PROCESS OLD FILE
    ProcessOldWorkbook wbNew, wbOld, wbGenerator
    
    wbNew.Close False
    wbOld.Close False

    MsgBox "The Macro Run Successfully", vbInformation


ExitHandler:

    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = xlCalculationAutomatic

Exit Sub

ErrHandler:
    MsgBox Err.Description, vbCritical
    Resume ExitHandler

End Sub