Option Explicit

'========================================================
' Lazily-created cache of "what was in the Comments cell
' just before this edit". Keyed by "SheetName|CellAddress".
' Used instead of Application.Undo (which was unreliable).
'========================================================

Private mPrevValues As Object

Private Function PrevValues() As Object
    If mPrevValues Is Nothing Then
        Set mPrevValues = CreateObject("Scripting.Dictionary")
    End If
    Set PrevValues = mPrevValues
End Function


'========================================================
' Excel resets "UserInterfaceOnly" protection every time
' the workbook is closed, so re-apply it on open. Without
' this, the toggle macro below would stop being able to
' write to a protected Comments column after a reopen.
'========================================================

Private Sub Workbook_Open()

    Dim ws As Worksheet

    On Error Resume Next
    For Each ws In ThisWorkbook.Worksheets
        If IS_DAILY_SHEET(ws.Name) Then
            If ws.ProtectContents Then
                ws.Protect UserInterfaceOnly:=True, DrawingObjects:=True, _
                           Contents:=True, Scenarios:=True, AllowFiltering:=True
            End If
        End If
    Next ws
    On Error GoTo 0

End Sub


'========================================================
' Cache the Comments cell's value the moment the user
' lands on it -- i.e. BEFORE the dropdown overwrites it.
'========================================================

Private Sub Workbook_SheetSelectionChange(ByVal Sh As Object, ByVal Target As Range)

    Dim commentsCol As Long
    Dim key As String

    If Target.Cells.CountLarge > 1 Then Exit Sub
    If Not IS_DAILY_SHEET(Sh.Name) Then Exit Sub

    commentsCol = FIND_HEADER(Sh, COMMENTS_HEADER)
    If commentsCol = 0 Then Exit Sub
    If Target.Column <> commentsCol Then Exit Sub
    If Target.Row <= HEADER_ROW Then Exit Sub

    key = Sh.Name & "|" & Target.Address(False, False)
    PrevValues()(key) = Trim(CStr(Target.value))

End Sub


'========================================================
' Multi-select toggle:
'   - selecting a new tag appends it ("a, b")
'   - selecting a tag that's already present removes it
'========================================================

Private Sub Workbook_SheetChange(ByVal Sh As Object, ByVal Target As Range)

    Dim commentsCol As Long
    Dim key As String

    Dim oldValue As String
    Dim newValue As String

    Dim values As Variant
    Dim item As String
    Dim result As String

    Dim i As Long
    Dim found As Boolean

    If Target.Cells.CountLarge > 1 Then Exit Sub
    If Not IS_DAILY_SHEET(Sh.Name) Then Exit Sub

    commentsCol = FIND_HEADER(Sh, COMMENTS_HEADER)
    If commentsCol = 0 Then Exit Sub
    If Target.Column <> commentsCol Then Exit Sub
    If Target.Row <= HEADER_ROW Then Exit Sub

    key = Sh.Name & "|" & Target.Address(False, False)
    newValue = Trim(CStr(Target.value))

    If Len(newValue) = 0 Then
        If PrevValues().Exists(key) Then PrevValues().Remove key
        Exit Sub
    End If

    If PrevValues().Exists(key) Then
        oldValue = CStr(PrevValues()(key))
    Else
        oldValue = ""
    End If

    On Error GoTo ExitHandler
    Application.EnableEvents = False

    If oldValue = "" Then
        Target.value = newValue
        GoTo Cleanup
    End If

    values = Split(oldValue, ",")
    result = ""
    found = False

    For i = LBound(values) To UBound(values)

        item = Trim(CStr(values(i)))

        If Len(item) = 0 Then
            ' skip stray blanks from double commas

        ElseIf StrComp(item, newValue, vbTextCompare) = 0 Then
            found = True

        Else
            If result = "" Then
                result = item
            Else
                result = result & ", " & item
            End If
        End If

    Next i

    If found Then
        Target.value = result
    Else
        If result = "" Then
            Target.value = newValue
        Else
            Target.value = result & ", " & newValue
        End If
    End If

Cleanup:
    key = Sh.Name & "|" & Target.Address(False, False)
    PrevValues()(key) = Trim(CStr(Target.value))

ExitHandler:
    Application.EnableEvents = True

End Sub
