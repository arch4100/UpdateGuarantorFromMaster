Attribute VB_Name = "Module15"
Option Explicit

'========================
' CONFIG (adjust if needed)
'========================
Private Const MASTER_PATH As String = "S:\Customer Engagement Division\PFD CLN\03.02_Bond Master LOI - Master Admin\.Master LOI - Master Admin.xlsm"
Private Const MASTER_SHEET As String = "List of Guarantors"

'Target file columns (ADJUST IF YOUR POLICY NO IS NOT HERE)
Private Const COL_TARGET_UEN As String = "C"
Private Const COL_TARGET_POLICYDATE As String = "K"
Private Const COL_TARGET_SUBCLASS As String = "E"
Private Const COL_TARGET_INTERMEDIARY As String = "H"

'Master columns
Private Const COL_MASTER_UEN As String = "G"
Private Const COL_MASTER_INTERMEDIARY As String = "B"
Private Const COL_MASTER_INDEMNITY As String = "C"
Private Const COL_MASTER_INSUREDNAME As String = "A"
Private Const COL_MASTER_UPLOADHIST As String = "X" 'NO/YES (auto flip NO->YES)

'Master LOI PDF folder
Private Const MASTER_LOI_PDF_FOLDER As String = "S:\Customer Engagement Division\PFD CLN\03.02_Bond Master LOI - Master Admin"

'Where to save the filtered draft attachment
Private Const EMAIL_EXCEL_DRAFT_FOLDER As String = "S:\Customer Engagement Division\PFD CLN\03.04_Bond LOI Return File Upload\zFiling\Email Excel Draft"

'File Upload Deposit folder
Private Const FINAL_OUTPUT_CSV As String = "S:\Common\Customer Engagement Division\Bond LOI Management\Upload Guarantor Details\GUARANTOR_LOI_PROCESS.csv"

'Subclass scope
Private Const SUBCLASS_1 As String = "BDFWIM"
Private Const SUBCLASS_2 As String = "FTFWOR"


Sub UpdateGuarantorFromMaster()

    Dim masterWB As Workbook, targetWB As Workbook
    Dim masterWS As Worksheet, targetWS As Worksheet
    Dim targetFile As Variant
    Dim lastRowMaster As Long, lastRowTarget As Long
    Dim i As Long, j As Long

    Dim uen As String
    Dim policyDate As Date
    Dim bestMatchRow As Long
    Dim latestValidDate As Date
    Dim rowFound As Boolean
    Dim checkVal As String

    Dim tgtInter As String, mstInter As String
    Dim indemnityDate As Date

    Dim processedMasterRows As Object
    Set processedMasterRows = CreateObject("Scripting.Dictionary")
    processedMasterRows.CompareMode = vbTextCompare

    Dim pendingEmails As Object
    Set pendingEmails = CreateObject("Scripting.Dictionary")
    pendingEmails.CompareMode = vbTextCompare

    Dim firstTimeList As Object
    Set firstTimeList = CreateObject("System.Collections.ArrayList")

    Dim masterChanged As Boolean
    masterChanged = False

    On Error GoTo EH

    'OPEN MASTER
    On Error Resume Next
    Set masterWB = Workbooks.Open(MASTER_PATH)
    On Error GoTo 0
    If masterWB Is Nothing Then
        MsgBox "Failed to open Master LOI workbook.", vbCritical
        Exit Sub
    End If

    Set masterWS = masterWB.Sheets(MASTER_SHEET)
    lastRowMaster = masterWS.Cells(masterWS.Rows.Count, COL_MASTER_UEN).End(xlUp).Row

    'PICK TARGET
    targetFile = Application.GetOpenFilename("Excel or CSV Files (*.xlsx;*.csv), *.xlsx;*.csv", , "Select Actual Data File")
    If targetFile = False Then
        masterWB.Close False
        Exit Sub
    End If

    If LCase$(Right$(CStr(targetFile), 4)) = ".csv" Then
        Workbooks.OpenText fileName:=CStr(targetFile), DataType:=xlDelimited, Comma:=True
        Set targetWB = ActiveWorkbook
    Else
        Set targetWB = Workbooks.Open(CStr(targetFile))
    End If

    Set targetWS = targetWB.Sheets(1)

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    lastRowTarget = targetWS.Cells(targetWS.Rows.Count, COL_TARGET_UEN).End(xlUp).Row

    '========================
    ' BEGIN TARGET LOOP
    '========================
    For i = 2 To lastRowTarget

        checkVal = Trim$(UCase$(CStr(targetWS.Cells(i, COL_TARGET_SUBCLASS).Value)))
        If checkVal <> SUBCLASS_1 And checkVal <> SUBCLASS_2 Then GoTo SkipToNext

        uen = Trim$(CStr(targetWS.Cells(i, COL_TARGET_UEN).Value))
        If uen = "" Then GoTo SkipToNext

        If IsDate(targetWS.Cells(i, COL_TARGET_POLICYDATE).Value) Then
            policyDate = CDate(targetWS.Cells(i, COL_TARGET_POLICYDATE).Value)
        Else
            GoTo SkipToNext
        End If

        tgtInter = NormalizeIntermediary(CStr(targetWS.Cells(i, COL_TARGET_INTERMEDIARY).Value))

        latestValidDate = 0
        rowFound = False
        bestMatchRow = 0

        'Find best match in master for same UEN + Intermediary, latest indemnity <= policyDate
        For j = 2 To lastRowMaster

            If Trim$(CStr(masterWS.Cells(j, COL_MASTER_UEN).Value)) = uen Then

                mstInter = NormalizeIntermediary(CStr(masterWS.Cells(j, COL_MASTER_INTERMEDIARY).Value))
                If mstInter = tgtInter Then

                    If IsDate(masterWS.Cells(j, COL_MASTER_INDEMNITY).Value) Then
                        indemnityDate = CDate(masterWS.Cells(j, COL_MASTER_INDEMNITY).Value)

                        If indemnityDate <= policyDate Then
                            If indemnityDate > latestValidDate Then
                                latestValidDate = indemnityDate
                                bestMatchRow = j
                                rowFound = True
                            End If
                        End If
                    End If

                End If
            End If

        Next j

        If rowFound Then

            'Update target mapping (as per your current mapping)
            ApplyGuarantorMapping targetWS, i, masterWS, bestMatchRow

            'AUTO: if Master X is NO => flip to YES and queue email (no prompt)
            If Trim$(UCase$(CStr(masterWS.Cells(bestMatchRow, COL_MASTER_UPLOADHIST).Value))) = "NO" Then

                If Not processedMasterRows.Exists(CStr(bestMatchRow)) Then
                    processedMasterRows.Add CStr(bestMatchRow), True

                    Dim insuredName As String
                    insuredName = CStr(masterWS.Cells(bestMatchRow, COL_MASTER_INSUREDNAME).Value)

                    'Flip Master X -> YES
                    masterWS.Cells(bestMatchRow, COL_MASTER_UPLOADHIST).Value = "YES"
                    masterChanged = True

                    'Track list for final msg (after all drafts)
                    If Len(Trim$(insuredName)) > 0 Then firstTimeList.Add insuredName

                    'Queue email: include masterRow so we can locate PDF later
                    Dim k As String
                    k = insuredName & "|" & uen & "|" & tgtInter & "|" & Format(latestValidDate, "yyyymmdd")
                    If Not pendingEmails.Exists(k) Then
                        pendingEmails.Add k, Array(insuredName, uen, tgtInter, latestValidDate, CLng(bestMatchRow))
                    End If
                End If

            End If

        End If

SkipToNext:
    Next i

    '========================
    ' AFTER ALL UPDATES: create drafts
    '========================
    Dim key As Variant, arr As Variant
    For Each key In pendingEmails.Keys

        arr = pendingEmails(key)
        'arr = Array(insuredName, uen, intermediary, indemnityDate, masterRow)

        Dim insured2 As String, uen2 As String, inter2 As String
        Dim indDate2 As Date, masterRow2 As Long
        Dim attachCsv As String, masterPdf As String

        insured2 = CStr(arr(0))
        uen2 = CStr(arr(1))
        inter2 = CStr(arr(2))
        indDate2 = CDate(arr(3))
        masterRow2 = CLng(arr(4))

        'Let your email macro generate filtered CSV if you pass blank,
        'BUT you asked to do it after updates  keep it explicit here (optional).
        attachCsv = "" 'leave blank to let PrepareFileUploadCheckEmailDraft build it

        masterPdf = FindMasterLOIPDF(masterWS, masterRow2)

        Call PrepareFileUploadCheckEmailDraft(insured2, uen2, inter2, indDate2, attachCsv, masterPdf)

    Next key

    'FINAL sorting/cleanup (keep your existing FinalizeTargetSortingAndCleanup if already in your module)
    FinalizeTargetSortingAndCleanup targetWS
    SaveTargetAsProcessCSV targetWB, targetWS

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents = True

    'Show list AFTER all drafts prepared
    If firstTimeList.Count > 0 Then
        firstTimeList.Sort
        MsgBox "First-time uploaded (Master Column X flipped from NO to YES):" & vbCrLf & _
               Join(firstTimeList.ToArray(), vbCrLf), vbInformation
    Else
        MsgBox "Update completed successfully. No first-time uploads detected.", vbInformation
    End If

    If Not masterWB Is Nothing Then masterWB.Close SaveChanges:=masterChanged
    Exit Sub

EH:
    Resume Cleanup

End Sub

'==================================================
' Find Master LOI PDF (INSURED - UEN - dd mmm yyyy.pdf)
' 1) Try exact date = Master indemnity date
' 2) Fallback: newest PDF by prefix "INSURED - UEN - "
'==================================================
Private Function FindMasterLOIPDF(ByVal masterWS As Worksheet, ByVal masterRow As Long) As String
    On Error GoTo EH

    Dim insuredName As String, uen As String
    Dim indDate As Date

    insuredName = Trim$(CStr(masterWS.Cells(masterRow, COL_MASTER_INSUREDNAME).Value))
    uen = Trim$(CStr(masterWS.Cells(masterRow, COL_MASTER_UEN).Value))

    If Not IsDate(masterWS.Cells(masterRow, COL_MASTER_INDEMNITY).Value) Then
        FindMasterLOIPDF = ""
        Exit Function
    End If
    indDate = CDate(masterWS.Cells(masterRow, COL_MASTER_INDEMNITY).Value)

    Dim p1 As String, p2 As String
    p1 = MASTER_LOI_PDF_FOLDER & "\" & insuredName & " - " & uen & " - " & Format(indDate, "dd mmm yyyy") & ".pdf"
    p2 = MASTER_LOI_PDF_FOLDER & "\" & insuredName & " - " & uen & " - " & Format(indDate, "d mmm yyyy") & ".pdf"

    If Len(Dir$(p1)) > 0 Then
        FindMasterLOIPDF = p1
        Exit Function
    End If
    If Len(Dir$(p2)) > 0 Then
        FindMasterLOIPDF = p2
        Exit Function
    End If

    'Fallback: newest by prefix
    Dim fso As Object, folder As Object, f As Object
    Dim prefix As String, bestPath As String
    Dim bestDT As Date: bestDT = 0

    prefix = insuredName & " - " & uen & " - "

    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(MASTER_LOI_PDF_FOLDER) Then
        FindMasterLOIPDF = ""
        Exit Function
    End If

    Set folder = fso.GetFolder(MASTER_LOI_PDF_FOLDER)

    For Each f In folder.Files
        If LCase$(fso.GetExtensionName(f.name)) = "pdf" Then
            If UCase$(Left$(f.name, Len(prefix))) = UCase$(prefix) Then
                If f.DateLastModified > bestDT Then
                    bestDT = f.DateLastModified
                    bestPath = f.Path
                End If
            End If
        End If
    Next f

    FindMasterLOIPDF = bestPath
    Exit Function

EH:
    FindMasterLOIPDF = ""
End Function

'==================================================
' POPUP HANDLER (UPDATED: QUEUE EMAIL INSTEAD OF CREATE IMMEDIATELY)
'==================================================
Private Sub HandleFirstTimeUploadPrompt( _
    ByVal masterWS As Worksheet, _
    ByVal masterRow As Long, _
    ByRef masterChanged As Boolean, _
    ByVal needCheckList As Object, _
    ByVal pendingEmails As Object, _
    ByVal uen As String, _
    ByVal intermediary As String, _
    ByVal policyNo As String, _
    ByVal indemnityDate As Date _
)

    Dim insuredName As String
    insuredName = CStr(masterWS.Cells(masterRow, COL_MASTER_INSUREDNAME).Value)

    Dim msg1 As String
    msg1 = "First-time file upload detected (Master Upload History = NO)." & vbCrLf & vbCrLf & _
           "Insured: " & insuredName & vbCrLf & _
           "Where to update: '" & MASTER_SHEET & "'!Column " & COL_MASTER_UPLOADHIST & " (Row " & masterRow & ")" & vbCrLf & vbCrLf & _
           "Do you want to update Upload History to YES?" & vbCrLf & _
           "Click YES = update Upload History to YES" & vbCrLf & _
           "Click NO = do not update"

    If MsgBox(msg1, vbQuestion + vbYesNo, "Confirm Upload History") = vbYes Then
        masterWS.Cells(masterRow, COL_MASTER_UPLOADHIST).Value = "YES"
        masterChanged = True
    Else
        'User said No: store for summary
        If insuredName <> "" Then needCheckList.Add insuredName

        Dim msg2 As String
        msg2 = "Do you want to prepare an Outlook email draft now?" & vbCrLf & vbCrLf & _
               "Click YES = create email draft (after macro finishes updating)" & vbCrLf & _
               "Click NO = skip"

        If MsgBox(msg2, vbQuestion + vbYesNo, "Prepare Email Draft") = vbYes Then

            'Queue (key prevents duplicates)
            Dim k As String
            k = insuredName & "|" & uen & "|" & intermediary & "|" & Format(indemnityDate, "yyyymmdd")

            If Not pendingEmails.Exists(k) Then
                pendingEmails.Add k, Array(insuredName, uen, intermediary, indemnityDate, CLng(masterRow))
            End If

        End If
    End If

End Sub

'==================================================
' CREATE FILTERED DRAFT ATTACHMENT FROM UPDATED TARGET SHEET
' - Filters on UEN + Intermediary (normalized) + PolicyDate >= indemnityDate
' - Subclass in (BDFWIM, FTFWOR)
' - Copies EXACT columns A:AE
' - Saves as XLSX in EMAIL_EXCEL_DRAFT_FOLDER
'==================================================
Private Function CreateFilteredExcelForChecking_FromSheet_AtoAE( _
    ByVal ws As Worksheet, _
    ByVal insuredName As String, _
    ByVal uen As String, _
    ByVal intermediary As String, _
    ByVal indemnityDate As Date _
) As String

    On Error GoTo EH

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TARGET_UEN).End(xlUp).Row
    If lastRow < 2 Then
        CreateFilteredExcelForChecking_FromSheet_AtoAE = ""
        Exit Function
    End If

    'FORCE copy A:AE
    Dim lastCol As Long
    lastCol = ws.Range("AE1").Column

    Dim normTgtInter As String
    normTgtInter = NormalizeIntermediary(CStr(intermediary))

    Dim newWB As Workbook, newWS As Worksheet
    Set newWB = Workbooks.Add(xlWBATWorksheet)
    Set newWS = newWB.Worksheets(1)

    'Header row A:AE
    ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol)).Copy Destination:=newWS.Cells(1, 1)

    Dim outRow As Long: outRow = 2

    Dim i As Long
    Dim vUEN As String, vSubclass As String, vInter As String
    Dim vPolDate As Date

    For i = 2 To lastRow

        vUEN = Trim$(CStr(ws.Cells(i, COL_TARGET_UEN).Value))
        If Len(vUEN) = 0 Then GoTo NextI
        If StrComp(vUEN, uen, vbTextCompare) <> 0 Then GoTo NextI

        vSubclass = Trim$(UCase$(CStr(ws.Cells(i, COL_TARGET_SUBCLASS).Value)))
        If vSubclass <> SUBCLASS_1 And vSubclass <> SUBCLASS_2 Then GoTo NextI

        vInter = NormalizeIntermediary(CStr(ws.Cells(i, COL_TARGET_INTERMEDIARY).Value))
        If StrComp(vInter, normTgtInter, vbTextCompare) <> 0 Then GoTo NextI

        If Not IsDate(ws.Cells(i, COL_TARGET_POLICYDATE).Value) Then GoTo NextI
        vPolDate = CDate(ws.Cells(i, COL_TARGET_POLICYDATE).Value)
        If vPolDate < indemnityDate Then GoTo NextI

        'Copy UPDATED row A:AE
        ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy Destination:=newWS.Cells(outRow, 1)
        outRow = outRow + 1

NextI:
    Next i

    If outRow = 2 Then
        newWB.Close SaveChanges:=False
        CreateFilteredExcelForChecking_FromSheet_AtoAE = ""
        Exit Function
    End If

    EnsureFolderExists EMAIL_EXCEL_DRAFT_FOLDER

    Dim fileName As String, fullPath As String
    fileName = SanitizeFileName(insuredName) & " - " & _
               SanitizeFileName(uen) & " - " & _
               SanitizeFileName(intermediary) & " - " & _
               Format(indemnityDate, "ddmmyyyy") & ".xlsx"

    fullPath = EMAIL_EXCEL_DRAFT_FOLDER & "\" & fileName

    Application.DisplayAlerts = False
    If Len(Dir$(fullPath)) > 0 Then Kill fullPath
    newWB.SaveAs fileName:=fullPath, FileFormat:=xlOpenXMLWorkbook
    newWB.Close SaveChanges:=False
    Application.DisplayAlerts = True

    CreateFilteredExcelForChecking_FromSheet_AtoAE = fullPath
    Exit Function

EH:
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not newWB Is Nothing Then newWB.Close SaveChanges:=False
    CreateFilteredExcelForChecking_FromSheet_AtoAE = ""
End Function

'==================================================
' FINALIZER: FILTER + SORT + DELETE BLANKS + SORTS
'==================================================
Private Sub FinalizeTargetSortingAndCleanup(ByVal ws As Worksheet)
    On Error GoTo SafeExit

    Dim lastRow As Long, lastCol As Long
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then GoTo SafeExit

    'Clear existing filter (if any)
    If ws.AutoFilterMode Then
        On Error Resume Next
        ws.ShowAllData
        On Error GoTo 0
        ws.AutoFilterMode = False
    End If

    'Apply filter on row 1 across used columns
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter

    '1) Sort Column U A-Z
    SortSheetByColumn ws, "U", lastRow, lastCol

    '2) Delete rows with blank in Column U (keep header)
    DeleteRowsWhereColumnBlank ws, "U", lastRow, lastCol

    'Recompute lastRow after deletion
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 2 Then GoTo SafeExit

    '3) Sort Column R A-Z
    SortSheetByColumn ws, "R", lastRow, lastCol

    '4) Sort Column B A-Z
    SortSheetByColumn ws, "B", lastRow, lastCol

SafeExit:
    'Optional: keep filter on (you asked to apply filter row 1)
End Sub

Private Sub SortSheetByColumn(ByVal ws As Worksheet, ByVal sortColLetter As String, ByVal lastRow As Long, ByVal lastCol As Long)
    With ws.Sort
        .SortFields.Clear
        .SortFields.Add key:=ws.Range(sortColLetter & "2:" & sortColLetter & lastRow), _
                        SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
        .SetRange ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With
End Sub

Private Sub DeleteRowsWhereColumnBlank(ByVal ws As Worksheet, ByVal colLetter As String, ByVal lastRow As Long, ByVal lastCol As Long)
    Dim rngData As Range, rngBlanks As Range
    Set rngData = ws.Range(ws.Cells(2, 1), ws.Cells(lastRow, lastCol))

    'Filter blanks in the specified column, then delete visible rows
    ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol)).AutoFilter Field:=ColumnLetterToNumber(colLetter), Criteria1:="="

    On Error Resume Next
    Set rngBlanks = rngData.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not rngBlanks Is Nothing Then
        rngBlanks.EntireRow.Delete
    End If

    'Clear filter criteria (keep filter row 1 active)
    On Error Resume Next
    ws.ShowAllData
    On Error GoTo 0
End Sub

Private Function ColumnLetterToNumber(ByVal colLetter As String) As Long
    ColumnLetterToNumber = Range(UCase$(colLetter) & "1").Column
End Function

'==================================================
' NORMALISE INTERMEDIARY FUNCTION (your logic)
'==================================================
Function NormalizeIntermediary(name As String) As String
    Dim t As String
    t = UCase$(Trim$(Replace(name, ".", "")))

    If t = "PREMIER INSURANCE AGENCIES PTE LTD" Then
        t = "HOWDEN PREMIER"
    ElseIf t = "INSURANCE SOLUTIONS PTE LTD" Then
        t = "INS-SOLUTIONS AGENCY PTE LTD"
    End If

    NormalizeIntermediary = t
End Function

Function MapGuarantorRole(role As String) As String
    Select Case UCase$(Trim$(role))
        Case "DIRECTOR": MapGuarantorRole = "D"
        Case "THIRD PARTY": MapGuarantorRole = "T"
        Case "SOLE PROPRIETOR": MapGuarantorRole = "S"
        Case "PARTNERSHIP": MapGuarantorRole = "P"
        Case Else: MapGuarantorRole = ""
    End Select
End Function

'==================================================
' MAPPING HELPERS
'==================================================
Private Sub ApplyGuarantorMapping( _
    ByVal targetWS As Worksheet, _
    ByVal targetRow As Long, _
    ByVal masterWS As Worksheet, _
    ByVal masterRow As Long _
)
    Dim mappings As Variant
    mappings = Array( _
        Array("U", "J", False), _
        Array("V", "K", True), _
        Array("W", "M", False), _
        Array("X", "N", True), _
        Array("Y", "P", False), _
        Array("Z", "Q", True), _
        Array("AA", "S", False), _
        Array("AB", "T", True), _
        Array("AC", "V", False), _
        Array("AD", "W", True) _
    )

    Dim i As Long
    For i = LBound(mappings) To UBound(mappings)
        WriteMappedValue targetWS, targetRow, masterWS, masterRow, mappings(i)
    Next i
End Sub

Private Sub WriteMappedValue( _
    ByVal targetWS As Worksheet, _
    ByVal targetRow As Long, _
    ByVal masterWS As Worksheet, _
    ByVal masterRow As Long, _
    ByVal mapping As Variant _
)
    Dim targetCol As String
    Dim sourceCol As String
    Dim isRole As Boolean

    targetCol = CStr(mapping(0))
    sourceCol = CStr(mapping(1))
    isRole = CBool(mapping(2))

    If isRole Then
        targetWS.Cells(targetRow, targetCol).Value = MapGuarantorRole(CStr(masterWS.Cells(masterRow, sourceCol).Value))
    Else
        targetWS.Cells(targetRow, targetCol).Value = masterWS.Cells(masterRow, sourceCol).Value
    End If
End Sub

'==================================================
' Utilities
'==================================================
Private Sub EnsureFolderExists(ByVal folderPath As String)
    On Error GoTo EH
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")

    If Len(Trim$(folderPath)) = 0 Then Exit Sub
    If fso.FolderExists(folderPath) Then Exit Sub

    Dim parentPath As String
    parentPath = fso.GetParentFolderName(folderPath)

    If Len(parentPath) > 0 And Not fso.FolderExists(parentPath) Then
        EnsureFolderExists parentPath
    End If

    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    Exit Sub
EH:
    'ignore
End Sub

Private Function SanitizeFileName(ByVal s As String) As String
    Dim badChars As Variant, i As Long
    badChars = Array("\", "/", ":", "*", "?", """", "<", ">", "|")
    For i = LBound(badChars) To UBound(badChars)
        s = Replace(s, CStr(badChars(i)), " ")
    Next i
    s = Trim$(s)
    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop
    SanitizeFileName = s
End Function
Private Sub SaveTargetAsProcessCSV(ByVal targetWB As Workbook, ByVal ws As Worksheet)
    On Error GoTo EH

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim folderPath As String
    folderPath = fso.GetParentFolderName(FINAL_OUTPUT_CSV)

    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath

    Application.DisplayAlerts = False

    'If file exists, remove it first (prevents SaveAs prompt / conflicts)
    If Len(Dir$(FINAL_OUTPUT_CSV)) > 0 Then
        On Error Resume Next
        Kill FINAL_OUTPUT_CSV
        On Error GoTo EH
    End If

    'Save as CSV (UTF-8 preferred)
    On Error Resume Next
    targetWB.SaveAs fileName:=FINAL_OUTPUT_CSV, FileFormat:=xlCSVUTF8, CreateBackup:=False
    If Err.Number <> 0 Then
        Err.Clear
        targetWB.SaveAs fileName:=FINAL_OUTPUT_CSV, FileFormat:=xlCSV, CreateBackup:=False
    End If
    On Error GoTo EH

    Application.DisplayAlerts = True
    Exit Sub

EH:
    Application.DisplayAlerts = True
    MsgBox "Failed to save output CSV to:" & vbCrLf & FINAL_OUTPUT_CSV & vbCrLf & vbCrLf & _
           "Error: " & Err.Description, vbExclamation
End Sub
