Option Explicit

'========================================================
' CONFIG
'========================================================

Public Const HEADER_ROW As Long = 3            ' row that has ROUTE NAME / Comments / etc.
Public Const PRODUCT_HEADER_ROW As Long = 2    ' row that has P / 4SQ / B / ... / TOTAL
Public Const COMMENTS_HEADER As String = "Comments"
Public Const LIST_SHEET As String = "CommentLists"
Public Const NAMED_RANGE As String = "CommentOptions"
Public Const LEGACY_CHECKED_HEADER As String = "Checked"
Public Const LEGACY_CHECKBOX_PREFIX As String = "ReturnsCheckbox_"

'========================================================
' MAIN SETUP
'========================================================

Public Sub SETUP_ALL_DAILY_SHEETS()

    Dim ws As Worksheet
    Dim processed As Long
    Dim failed As Long
    Dim errorLog As String
    Dim oldEvents As Boolean
    Dim oldScreen As Boolean
    Dim nameExists As Boolean

    oldEvents = Application.EnableEvents
    oldScreen = Application.ScreenUpdating

    On Error Resume Next
    If ActiveWindow.SelectedSheets.Count > 1 Then ActiveSheet.Select
    On Error GoTo FatalError

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    'Build the (hardcoded) Comments dropdown options
    CREATE_COMMENT_OPTIONS

    nameExists = False
    On Error Resume Next
    nameExists = (Len(ThisWorkbook.Names(NAMED_RANGE).Name) > 0)
    On Error GoTo 0

    If Not nameExists Then
        MsgBox "Setup stopped: the '" & NAMED_RANGE & "' dropdown list could not be built.", _
               vbCritical, "Daily Returns"
        GoTo Cleanup
    End If

    'Process every daily sheet (1ST - 31ST)
    For Each ws In ThisWorkbook.Worksheets

        If IS_DAILY_SHEET(ws.Name) Then

            On Error GoTo SheetError

            ws.Select
            SETUP_COMMENTS ws

            processed = processed + 1

        End If

ContinueLoop:
        On Error GoTo FatalError

    Next ws

Cleanup:

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox _
        "SETUP COMPLETE" & vbCrLf & vbCrLf & _
        "Daily sheets processed: " & processed & vbCrLf & _
        "Sheets with errors: " & failed & IIf(Len(errorLog) > 0, vbCrLf & errorLog, ""), _
        vbInformation, _
        "Daily Returns"

    Exit Sub

SheetError:

    failed = failed + 1
    errorLog = errorLog & vbCrLf & ws.Name & ": " & Err.Number & " - " & Err.Description

    Debug.Print "SHEET ERROR: " & ws.Name & " -- " & Err.Number & " - " & Err.Description

    Err.Clear
    Resume ContinueLoop

FatalError:

    Application.EnableEvents = oldEvents
    Application.ScreenUpdating = oldScreen

    MsgBox _
        "Error " & Err.Number & vbCrLf & Err.Description, _
        vbCritical, "Daily Returns Error"

End Sub


'========================================================
' IDENTIFY DAILY SHEETS
'========================================================

Public Function IS_DAILY_SHEET(ByVal sheetName As String) As Boolean

    Dim s As String
    s = UCase(Trim(sheetName))

    Select Case s
        Case "1ST", "2ND", "3RD", _
             "4TH", "5TH", "6TH", "7TH", "8TH", "9TH", _
             "10TH", "11TH", "12TH", "13TH", "14TH", _
             "15TH", "16TH", "17TH", "18TH", "19TH", _
             "20TH", "21ST", "22ND", "23RD", "24TH", _
             "25TH", "26TH", "27TH", "28TH", "29TH", _
             "30TH", "31ST"
            IS_DAILY_SHEET = True
        Case Else
            IS_DAILY_SHEET = False
    End Select

End Function


'========================================================
' HARDCODED DROPDOWN OPTIONS
' (A named-range/list-sheet source is still required even
'  though the list is hardcoded, because Excel's Data
'  Validation "Formula1" literal-list field is capped at
'  255 characters and this list is ~630 characters.)
'========================================================

Public Sub CREATE_COMMENT_OPTIONS()

    Dim ws As Worksheet
    Dim options As Variant
    Dim i As Long
    Dim r As Long

    Dim opts(1 To 70) As String
    
    opts(1) = "p_oversupply"
    opts(2) = "p_wrong_LPO"
    opts(3) = "p_paper_damage"
    opts(4) = "p_underbaked_and_undersize"
    opts(5) = "p_damage_by_sales_men"
    opts(6) = "p_short_expriry"
    opts(7) = "p_faded_dates"
    
    opts(8) = "s_oversupply"
    opts(9) = "s_wrong_LPO"
    opts(10) = "s_paper_damage"
    opts(11) = "s_underbaked_and_undersize"
    opts(12) = "s_damage_by_sales_men"
    opts(13) = "s_short_expriry"
    opts(14) = "s_faded_dates"
    
    opts(15) = "b_oversupply"
    opts(16) = "b_wrong_LPO"
    opts(17) = "b_paper_damage"
    opts(18) = "b_underbaked_and_undersize"
    opts(19) = "b_damage_by_sales_men"
    opts(20) = "b_short_expriry"
    opts(21) = "b_faded_dates"
    
    opts(22) = "200g_oversupply"
    opts(23) = "200g_wrong_LPO"
    opts(24) = "200g_paper_damage"
    opts(25) = "200g_underbaked_and_undersize"
    opts(26) = "200g_damage_by_sales_men"
    opts(27) = "200g_short_expriry"
    opts(28) = "200g_faded_dates"
    
    opts(29) = "op_oversupply"
    opts(30) = "op_wrong_LPO"
    opts(31) = "op_paper_damage"
    opts(32) = "op_underbaked_and_undersize"
    opts(33) = "op_damage_by_sales_men"
    opts(34) = "op_short_expriry"
    opts(35) = "op_faded_dates"
    
    opts(36) = "800g_oversupply"
    opts(37) = "800g_wrong_LPO"
    opts(38) = "800g_paper_damage"
    opts(39) = "800g_underbaked_and_undersize"
    opts(40) = "800g_damage_by_sales_men"
    opts(41) = "800g_short_expriry"
    opts(42) = "800g_faded_dates"
    
    opts(43) = "600g_oversupply"
    opts(44) = "600g_wrong_LPO"
    opts(45) = "600g_paper_damage"
    opts(46) = "600g_underbaked_and_undersize"
    opts(47) = "600g_damage_by_sales_men"
    opts(48) = "600g_short_expriry"
    opts(49) = "600g_faded_dates"
    
    opts(50) = "4_square_oversupply"
    opts(51) = "4_square_wrong_LPO"
    opts(52) = "4_square_paper_damage"
    opts(53) = "4_square_underbaked_and_undersize"
    opts(54) = "4_square_damage_by_sales_men"
    opts(55) = "4_square_short_expriry"
    opts(56) = "4_square_faded_dates"
    
    opts(57) = "xtra_loong_oversupply"
    opts(58) = "xtra_loong_wrong_LPO"
    opts(59) = "xtra_loong_paper_damage"
    opts(60) = "xtra_loong_underbaked_and_undersize"
    opts(61) = "xtra_loong_damage_by_sales_men"
    opts(62) = "xtra_loong_short_expriry"
    opts(63) = "xtra_loong_faded_dates"
    
    opts(64) = "bb_oversupply"
    opts(65) = "bb_wrong_LPO"
    opts(66) = "bb_paper_damage"
    opts(67) = "bb_underbaked_and_undersize"
    opts(68) = "bb_damage_by_sales_men"
    opts(69) = "bb_short_expriry"
    opts(70) = "bb_faded_dates"

    options = opts

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(LIST_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.Name = LIST_SHEET
    End If

    ws.Visible = xlSheetVisible
    ws.Cells.Clear
    ws.Cells(1, 1).value = "COMMENTS DROPDOWN OPTIONS"

    r = 2
    For i = LBound(options) To UBound(options)
        ws.Cells(r, 1).value = options(i)
        r = r + 1
    Next i

    On Error Resume Next
    ThisWorkbook.Names(NAMED_RANGE).Delete
    On Error GoTo 0

    ThisWorkbook.Names.Add _
        Name:=NAMED_RANGE, _
        RefersTo:="='" & LIST_SHEET & "'!$A$2:$A$" & CStr(r - 1)

    ws.Visible = xlSheetVeryHidden

End Sub


'========================================================
' CREATE / RELOCATE COMMENTS COLUMN + DROPDOWN
' Forces the Comments column to sit immediately after
' TOTAL, deleting any stray "Comments" column elsewhere.
' Also handles sheet protection: a protected sheet hides
' the dropdown arrow on any LOCKED cell (even though the
' input message still shows), which is almost certainly
' why you were only seeing the tooltip and no arrow.
'========================================================

Public Sub SETUP_COMMENTS(ByVal ws As Worksheet)

    Dim totalCol As Long
    Dim expectedCol As Long
    Dim existingCol As Long
    Dim commentsCol As Long
    Dim lastRow As Long
    Dim targetRange As Range
    Dim guard As Long
    Dim wasProtected As Boolean

    '----------------------------------------------------
    ' Temporarily unprotect (if we can) so we can edit
    ' headers/validation/lock-state.
    '----------------------------------------------------

    wasProtected = ws.ProtectContents

    If wasProtected Then
        On Error Resume Next
        ws.Unprotect
        On Error GoTo 0

        If ws.ProtectContents Then
            MsgBox "Sheet '" & ws.Name & "' is protected with a password I don't have, " & _
                   "so I can't set up its Comments column. Please unprotect it (Review > " & _
                   "Unprotect Sheet) and run setup again.", vbExclamation, "Protected Sheet"
            Exit Sub
        End If
    End If

    guard = 0
    Do
        existingCol = FIND_HEADER(ws, COMMENTS_HEADER)

        totalCol = FIND_HEADER_IN_ROW(ws, PRODUCT_HEADER_ROW, "TOTAL")
        If totalCol = 0 Then totalCol = LAST_USED_COLUMN(ws)
        expectedCol = totalCol + 1

        If existingCol = 0 Or existingCol = expectedCol Then Exit Do

        ws.Columns(existingCol).Delete

        guard = guard + 1
        If guard > 10 Then Exit Do
    Loop

    commentsCol = expectedCol

    With ws.Cells(HEADER_ROW, commentsCol)
        .value = COMMENTS_HEADER
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
        .Locked = False
    End With

    ws.Columns(commentsCol).ColumnWidth = 40

    lastRow = LAST_USED_ROW(ws)

    If lastRow > HEADER_ROW Then

        Set targetRange = ws.Range( _
            ws.Cells(HEADER_ROW + 1, commentsCol), _
            ws.Cells(lastRow, commentsCol))

        ' Unlocked so the dropdown arrow will show once re-protected
        targetRange.Locked = False

        On Error Resume Next
        targetRange.Validation.Delete
        On Error GoTo 0

        On Error GoTo DropdownError

        With targetRange.Validation
            .Add _
                Type:=xlValidateList, _
                AlertStyle:=xlValidAlertStop, _
                Operator:=xlBetween, _
                Formula1:="=" & NAMED_RANGE

            .IgnoreBlank = True
            .InCellDropdown = True
            .ShowInput = True
            .ShowError = True

            .InputTitle = "Return Comment"
            .InputMessage = "Pick an issue. Pick it again to remove it."

            .ErrorTitle = "Invalid Option"
            .ErrorMessage = "Please select an option from the dropdown."
        End With

        On Error GoTo 0

    End If

    '----------------------------------------------------
    ' Re-protect if it was protected before -- using
    ' UserInterfaceOnly so our macros can still write to
    ' the sheet (e.g. the multi-select toggle) without
    ' needing to unprotect every time.
    '----------------------------------------------------

    If wasProtected Then
        ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, _
                   Contents:=True, Scenarios:=True, AllowFiltering:=True
    End If

    Exit Sub

DropdownError:

    Dim msg As String
    msg = "Could not add the dropdown on sheet '" & ws.Name & "'." & vbCrLf & _
          "Error " & Err.Number & ": " & Err.Description

    Debug.Print msg

    If wasProtected Then
        ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, _
                   Contents:=True, Scenarios:=True, AllowFiltering:=True
    End If

    MsgBox msg, vbExclamation, "Dropdown Error"

End Sub


'========================================================
' ONE-TIME CLEANUP: remove the old "Checked" column and
' any leftover checkbox shapes from a previous version.
'========================================================

Public Sub REMOVE_LEGACY_CHECKED_COLUMN()

    Dim ws As Worksheet
    Dim checkedCol As Long
    Dim i As Long
    Dim shp As Shape
    Dim cleaned As Long
    Dim wasProtected As Boolean

    Application.EnableEvents = False
    Application.ScreenUpdating = False

    For Each ws In ThisWorkbook.Worksheets

        If IS_DAILY_SHEET(ws.Name) Then

            wasProtected = ws.ProtectContents
            If wasProtected Then
                On Error Resume Next
                ws.Unprotect
                On Error GoTo 0
            End If

            On Error Resume Next
            For i = ws.Shapes.Count To 1 Step -1
                Set shp = ws.Shapes(i)
                If Not shp Is Nothing Then
                    If Left(shp.Name, Len(LEGACY_CHECKBOX_PREFIX)) = LEGACY_CHECKBOX_PREFIX Then
                        shp.Delete
                    End If
                End If
            Next i
            On Error GoTo 0

            checkedCol = FIND_HEADER(ws, LEGACY_CHECKED_HEADER)

            If checkedCol > 0 Then
                ws.Columns(checkedCol).Delete
                cleaned = cleaned + 1
            End If

            If wasProtected And Not ws.ProtectContents Then
                ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, _
                           Contents:=True, Scenarios:=True, AllowFiltering:=True
            End If

        End If

    Next ws

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "Removed the legacy 'Checked' column/checkboxes from " & _
           cleaned & " sheet(s).", vbInformation, "Cleanup"

End Sub


'========================================================
' ONE CLICK: clean up legacy artifacts AND rebuild
'========================================================

Public Sub CLEAN_AND_REBUILD_ALL()
    REMOVE_LEGACY_CHECKED_COLUMN
    SETUP_ALL_DAILY_SHEETS
End Sub


'========================================================
' DIAGNOSE -- run with a daily sheet active
'========================================================

Public Sub DIAGNOSE_SETUP()

    Dim ws As Worksheet
    Dim totalCol As Long, commentsCol As Long
    Dim namedRangeOK As Boolean
    Dim itemCount As Long
    Dim msg As String

    Set ws = ActiveSheet

    If Not IS_DAILY_SHEET(ws.Name) Then
        MsgBox "Select a daily sheet such as 1ST, 2ND or 3RD first.", vbExclamation, "Diagnose"
        Exit Sub
    End If

    totalCol = FIND_HEADER_IN_ROW(ws, PRODUCT_HEADER_ROW, "TOTAL")
    commentsCol = FIND_HEADER(ws, COMMENTS_HEADER)

    namedRangeOK = False
    itemCount = 0
    On Error Resume Next
    itemCount = ThisWorkbook.Names(NAMED_RANGE).RefersToRange.Cells.Count
    namedRangeOK = (Err.Number = 0)
    Err.Clear
    On Error GoTo 0

    msg = "Sheet: " & ws.Name & vbCrLf & _
          "Sheet protected: " & ws.ProtectContents & vbCrLf & _
          "'TOTAL' header found in column: " & IIf(totalCol = 0, "NOT FOUND", CStr(totalCol)) & vbCrLf & _
          "'Comments' header currently at column: " & IIf(commentsCol = 0, "NOT FOUND", CStr(commentsCol)) & vbCrLf & _
          IIf(commentsCol > 0, "Comments column locked: " & ws.Cells(HEADER_ROW + 1, commentsCol).Locked & vbCrLf, "") & _
          "'" & NAMED_RANGE & "' named range: " & IIf(namedRangeOK, CStr(itemCount) & " items", "MISSING/BROKEN")

    Debug.Print msg
    MsgBox msg, vbInformation, "Diagnose Setup"

End Sub


'========================================================
' FIND HEADER (defaults to row 3 for backward compatibility)
'========================================================

Public Function FIND_HEADER(ByVal ws As Worksheet, ByVal headerName As String) As Long
    FIND_HEADER = FIND_HEADER_IN_ROW(ws, HEADER_ROW, headerName)
End Function

Public Function FIND_HEADER_IN_ROW( _
    ByVal ws As Worksheet, _
    ByVal rowNum As Long, _
    ByVal headerName As String) As Long

    Dim c As Long
    Dim lastCol As Long

    lastCol = LAST_USED_COLUMN(ws)

    For c = 1 To lastCol
        If NORMALIZE_TEXT(ws.Cells(rowNum, c).value) = NORMALIZE_TEXT(headerName) Then
            FIND_HEADER_IN_ROW = c
            Exit Function
        End If
    Next c

    FIND_HEADER_IN_ROW = 0

End Function


'========================================================
' NORMALIZE TEXT
'========================================================

Private Function NORMALIZE_TEXT(ByVal value As Variant) As String

    Dim txt As String
    txt = CStr(value)
    txt = Replace(txt, Chr(160), " ")
    txt = Replace(txt, vbCr, " ")
    txt = Replace(txt, vbLf, " ")
    txt = Trim(txt)

    Do While InStr(txt, "  ") > 0
        txt = Replace(txt, "  ", " ")
    Loop

    NORMALIZE_TEXT = UCase(txt)

End Function


'========================================================
' LAST USED ROW / COLUMN
'========================================================

Public Function LAST_USED_ROW(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByRows, _
        SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LAST_USED_ROW = HEADER_ROW
    Else
        LAST_USED_ROW = lastCell.Row
        If LAST_USED_ROW < HEADER_ROW Then LAST_USED_ROW = HEADER_ROW
    End If

End Function

Public Function LAST_USED_COLUMN(ByVal ws As Worksheet) As Long

    Dim lastCell As Range

    On Error Resume Next
    Set lastCell = ws.Cells.Find( _
        What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
        LookAt:=xlPart, SearchOrder:=xlByColumns, _
        SearchDirection:=xlPrevious, MatchCase:=False)
    On Error GoTo 0

    If lastCell Is Nothing Then
        LAST_USED_COLUMN = 1
    Else
        LAST_USED_COLUMN = lastCell.Column
    End If

End Function


'========================================================
' REPAIR DROPDOWN ON CURRENT SHEET
'========================================================

Public Sub REPAIR_CURRENT_DROPDOWN()

    Dim ws As Worksheet
    Set ws = ActiveSheet

    If Not IS_DAILY_SHEET(ws.Name) Then
        MsgBox "Select a daily sheet such as 1ST, 2ND or 3RD.", _
            vbExclamation, "Daily Returns"
        Exit Sub
    End If

    Application.EnableEvents = False
    SETUP_COMMENTS ws
    Application.EnableEvents = True

    MsgBox "Dropdown repaired on " & ws.Name & ".", vbInformation, "Comments"

End Sub


'========================================================
' UNGROUP SHEETS
'========================================================

Public Sub UNGROUP_SHEETS()
    On Error Resume Next
    If ActiveWindow.SelectedSheets.Count > 1 Then ActiveSheet.Select
    On Error GoTo 0
End Sub


'========================================================
' RESET EXCEL
'========================================================

Public Sub RESET_EXCEL()
    On Error Resume Next
    If ActiveWindow.SelectedSheets.Count > 1 Then ActiveSheet.Select

    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Application.DisplayAlerts = True
    On Error GoTo 0

    MsgBox "Excel reset successfully.", vbInformation, "Reset"
End Sub

