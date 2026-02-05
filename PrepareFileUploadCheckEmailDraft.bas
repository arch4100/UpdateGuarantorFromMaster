Attribute VB_Name = "Module1"
Option Explicit

'========================
' EDIT THESE
'========================
Private Const MASTER_SOFTCOPY_LINK As String = "S:\Customer Engagement Division\PFD CLN\03.02_Bond Master LOI - Master Admin" 'your "here" link target

'Where to save the filtered draft attachment
Private Const EMAIL_EXCEL_DRAFT_FOLDER As String = "S:\Customer Engagement Division\PFD CLN\03.04_Bond LOI Return File Upload\zFiling\Email Excel Draft"

'Default recipients (optional)
Private Const DEFAULT_TO As String = "candy.law@sompo.com.sg"
Private Const DEFAULT_CC As String = "agnes.li@sompo.com.sg; alice.chong@sompo.com.sg; siewhong.see@sompo.com.sg"

'Target columns (must match your target file structure)
Private Const COL_TARGET_UEN As String = "C"
Private Const COL_TARGET_SUBCLASS As String = "E"
Private Const COL_TARGET_INTERMEDIARY As String = "H"
Private Const COL_TARGET_POLICYDATE As String = "K"

'Subclass scope
Private Const SUBCLASS_1 As String = "BDFWIM"
Private Const SUBCLASS_2 As String = "FTFWOR"

'========================
' MAIN MACRO (called by UpdateGuarantorFromMaster)
' - Attaches:
'   1) filtered csv (existing logic)
'   2) Master LOI PDF (NEW, optional, as second attachment)
'========================
Public Sub PrepareFileUploadCheckEmailDraft( _
    ByVal insuredName As String, _
    ByVal uen As String, _
    ByVal intermediary As String, _
    ByVal indemnityDate As Date, _
    Optional ByVal attachPath As String = "", _
    Optional ByVal masterPdfPath As String = "" _
)
    On Error GoTo EH

    Dim finalAttachPath As String
    finalAttachPath = attachPath

    If Len(Trim$(finalAttachPath)) = 0 Then
        finalAttachPath = CreateFilteredExcelForChecking(insuredName, uen, intermediary, indemnityDate)
    End If

    Dim olApp As Object, mail As Object
    Set olApp = GetOutlookApp()
    If olApp Is Nothing Then
        MsgBox "Outlook is not available.", vbExclamation
        Exit Sub
    End If

    Set mail = olApp.CreateItem(0) 'olMailItem

    Dim subj As String
    subj = "[FOR CHECKING] [FILE UPLOAD] " & insuredName & " - " & uen & " - " & intermediary & " - " & _
           Format(indemnityDate, "dd/mm/yyyy") & " " & Day(Date) & OrdSuffix(Day(Date)) & " " & Format(Date, "mmmm yyyy")

    Dim html As String
    html = BuildHtmlBody(insuredName, intermediary, indemnityDate)

    With mail
        If Len(DEFAULT_TO) > 0 Then .To = DEFAULT_TO
        If Len(DEFAULT_CC) > 0 Then .CC = DEFAULT_CC

        .Subject = subj
        .BodyFormat = 2 'olFormatHTML

        'Let Outlook insert signature
        .Display

        Dim existingSig As String
        existingSig = .HTMLBody
        existingSig = TrimLeadingSignatureBlanks(existingSig)

        'Prepend our content
        .HTMLBody = html & existingSig

        'Attachment #1 (filtered csv)
        If Len(Trim$(finalAttachPath)) > 0 Then
            If FileExists(finalAttachPath) Then .Attachments.Add finalAttachPath
        End If

        'Attachment #2 (Master LOI pdf)
        If Len(Trim$(masterPdfPath)) > 0 Then
            If FileExists(masterPdfPath) Then .Attachments.Add masterPdfPath
        End If
    End With

    Exit Sub

EH:
    MsgBox "Failed to prepare email draft: " & Err.Description, vbExclamation
End Sub

'==================================================
' CREATE FILTERED DRAFT ATTACHMENT (CSV)  (UNCHANGED)
'==================================================
Private Function CreateFilteredExcelForChecking( _
    ByVal insuredName As String, _
    ByVal uen As String, _
    ByVal intermediary As String, _
    ByVal indemnityDate As Date _
) As String

    On Error GoTo EH

    Dim ws As Worksheet
    Set ws = FindTargetWorksheetByUEN(uen)
    If ws Is Nothing Then
        CreateFilteredExcelForChecking = ""
        Exit Function
    End If

    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TARGET_UEN).End(xlUp).Row
    If lastRow < 2 Then
        CreateFilteredExcelForChecking = ""
        Exit Function
    End If

    Dim lastCol As Long
    lastCol = GetLastUsedColumn(ws)
    If lastCol < 1 Then lastCol = 1

    Dim normTgtInter As String
    normTgtInter = NormalizeIntermediary(CStr(intermediary)) 'must exist in your project

    Dim newWB As Workbook, newWS As Worksheet
    Set newWB = Workbooks.Add(xlWBATWorksheet)
    Set newWS = newWB.Worksheets(1)

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

        ws.Range(ws.Cells(i, 1), ws.Cells(i, lastCol)).Copy Destination:=newWS.Cells(outRow, 1)
        outRow = outRow + 1

NextI:
    Next i

    If outRow = 2 Then
        newWB.Close SaveChanges:=False
        CreateFilteredExcelForChecking = ""
        Exit Function
    End If

    Dim fileName As String, fullPath As String
    fileName = SanitizeFileName(insuredName) & " - " & _
               SanitizeFileName(uen) & " - " & _
               SanitizeFileName(intermediary) & " - " & _
               Format(indemnityDate, "ddmmyyyy") & ".csv"

    EnsureFolderExists EMAIL_EXCEL_DRAFT_FOLDER
    fullPath = EMAIL_EXCEL_DRAFT_FOLDER & "\" & fileName

    If FileExists(fullPath) Then
        On Error Resume Next
        Kill fullPath
        On Error GoTo EH
    End If

    Application.DisplayAlerts = False
    On Error Resume Next
    newWB.SaveAs fileName:=fullPath, FileFormat:=xlCSVUTF8
    If Err.Number <> 0 Then
        Err.Clear
        newWB.SaveAs fileName:=fullPath, FileFormat:=xlCSV
    End If
    On Error GoTo EH
    newWB.Close SaveChanges:=False
    Application.DisplayAlerts = True

    CreateFilteredExcelForChecking = fullPath
    Exit Function

EH:
    On Error Resume Next
    Application.DisplayAlerts = True
    If Not newWB Is Nothing Then newWB.Close SaveChanges:=False
    On Error GoTo 0
    CreateFilteredExcelForChecking = ""
End Function

'==================================================
' Find worksheet among open workbooks by searching for UEN
'==================================================
Private Function FindTargetWorksheetByUEN(ByVal uen As String) As Worksheet
    On Error GoTo EH

    Dim wb As Workbook, ws As Worksheet

    If Not ActiveWorkbook Is Nothing Then
        For Each ws In ActiveWorkbook.Worksheets
            If WorksheetHasUEN(ws, uen) Then
                Set FindTargetWorksheetByUEN = ws
                Exit Function
            End If
        Next ws
    End If

    For Each wb In Application.Workbooks
        For Each ws In wb.Worksheets
            If WorksheetHasUEN(ws, uen) Then
                Set FindTargetWorksheetByUEN = ws
                Exit Function
            End If
        Next ws
    Next wb

    Set FindTargetWorksheetByUEN = Nothing
    Exit Function

EH:
    Set FindTargetWorksheetByUEN = Nothing
End Function

Private Function WorksheetHasUEN(ByVal ws As Worksheet, ByVal uen As String) As Boolean
    On Error GoTo EH

    Dim lastRow As Long, i As Long
    lastRow = ws.Cells(ws.Rows.Count, COL_TARGET_UEN).End(xlUp).Row
    If lastRow < 2 Then
        WorksheetHasUEN = False
        Exit Function
    End If

    Dim scanTo As Long
    scanTo = lastRow
    If scanTo > 2000 Then scanTo = 2000

    For i = 2 To scanTo
        If StrComp(Trim$(CStr(ws.Cells(i, COL_TARGET_UEN).Value)), uen, vbTextCompare) = 0 Then
            WorksheetHasUEN = True
            Exit Function
        End If
    Next i

    If lastRow > 2000 Then
        For i = 2001 To lastRow
            If StrComp(Trim$(CStr(ws.Cells(i, COL_TARGET_UEN).Value)), uen, vbTextCompare) = 0 Then
                WorksheetHasUEN = True
                Exit Function
            End If
        Next i
    End If

    WorksheetHasUEN = False
    Exit Function

EH:
    WorksheetHasUEN = False
End Function

'========================
' HTML TEMPLATE
'========================
Private Function BuildHtmlBody(ByVal insuredName As String, ByVal intermediary As String, ByVal indemnityDate As Date) As String
    Dim t As String

    t = ""
    t = t & "<div style='font-family:Calibri,Arial; font-size:11pt;'>"
    t = t & "Good Day Candy,<br><br>"
    t = t & "Kindly assist to check <b>" & HtmlEncode(insuredName) & "</b> (First Upload).<br><br>"
    t = t & "Master LOI Softcopies can be found <a href='" & HtmlEncode(MASTER_SOFTCOPY_LINK) & "'>here</a>.<br><br>"

    t = t & "<table cellpadding='8' cellspacing='0' style='border-collapse:collapse; font-family:Calibri,Arial; font-size:11pt;'>"
    t = t & "  <tr>"
    t = t & "    <th style='border:1px solid #808080; background:transparent; color:#000000; text-align:center; vertical-align:middle;'>Name of Insured</th>"
    t = t & "    <th style='border:1px solid #808080; background:transparent; color:#000000; text-align:center; vertical-align:middle;'>Signed Intermediary</th>"
    t = t & "    <th style='border:1px solid #808080; background:transparent; color:#000000; text-align:center; vertical-align:middle;'>Indemnity Date</th>"
    t = t & "  </tr>"

    t = t & "  <tr>"
    t = t & "    <td style='border:1px solid #808080; background:transparent; color:#000000; text-align:center; vertical-align:middle;'><b>" & HtmlEncode(insuredName) & "</b></td>"
    t = t & "    <td style='border:1px solid #808080; background:transparent; color:#000000; text-align:center; vertical-align:middle;'><b>" & HtmlEncode(intermediary) & "</b></td>"
    t = t & "    <td style='border:1px solid #808080; background:transparent; color:#000000; text-align:center; vertical-align:middle;'><b>" & Format(indemnityDate, "dd/mm/yyyy") & "</b></td>"
    t = t & "  </tr>"
    t = t & "</table><br>"

    t = t & "Thank You!"
    t = t & "</div>"

    BuildHtmlBody = t
End Function

'========================
' Outlook App
'========================
Private Function GetOutlookApp() As Object
    On Error Resume Next
    Set GetOutlookApp = GetObject(, "Outlook.Application")
    If GetOutlookApp Is Nothing Then
        Set GetOutlookApp = CreateObject("Outlook.Application")
    End If
    On Error GoTo 0
End Function

Private Function TrimLeadingSignatureBlanks(ByVal html As String) As String
    Dim s As String
    s = html

    Dim re As Object
    Set re = CreateObject("VBScript.RegExp")
    re.Global = True
    re.ignoreCase = True
    re.Multiline = True

    Dim i As Long
    For i = 1 To 6
        re.pattern = "^\s*(<br[^>]*>\s*)+"
        s = re.Replace(s, "")

        re.pattern = "^\s*<p[^>]*>\s*(&nbsp;|\s|<br[^>]*>|<o:p>\s*</o:p>)*\s*</p>"
        s = re.Replace(s, "")

        re.pattern = "^\s*<div[^>]*>\s*(&nbsp;|\s|<br[^>]*>|<o:p>\s*</o:p>)*\s*</div>"
        s = re.Replace(s, "")
    Next i

    TrimLeadingSignatureBlanks = s
End Function

Private Function GetLastUsedColumn(ByVal ws As Worksheet) As Long
    On Error GoTo EH

    Dim lastCell As Range
    Set lastCell = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                                 LookIn:=xlFormulas, LookAt:=xlPart, _
                                 SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)

    If lastCell Is Nothing Then
        GetLastUsedColumn = 1
    Else
        GetLastUsedColumn = lastCell.Column
    End If
    Exit Function

EH:
    GetLastUsedColumn = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
End Function

Private Sub EnsureFolderExists(ByVal folderPath As String)
    On Error GoTo EH
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    If Not fso.FolderExists(folderPath) Then fso.CreateFolder folderPath
    Exit Sub
EH:
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

Private Function FileExists(ByVal p As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir$(p)) > 0)
    On Error GoTo 0
End Function

Private Function HtmlEncode(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    s = Replace(s, """", "&quot;")
    HtmlEncode = s
End Function

Private Function OrdSuffix(ByVal n As Long) As String
    Dim nMod100 As Long, nMod10 As Long
    nMod100 = n Mod 100
    nMod10 = n Mod 10

    If nMod100 >= 11 And nMod100 <= 13 Then
        OrdSuffix = "th"
    Else
        Select Case nMod10
            Case 1: OrdSuffix = "st"
            Case 2: OrdSuffix = "nd"
            Case 3: OrdSuffix = "rd"
            Case Else: OrdSuffix = "th"
        End Select
    End If
End Function
