Option Explicit

' --- refresh ---------------------------------------------------------------
' Refreshes each connection synchronously and reports the outcome to the
' caller. Returns "" when every query refreshed, or a one-line reason: the
' connection that failed, or that there was nothing to refresh at all.
'
' It must never MsgBox: Excel runs hidden, so a modal waits forever on an
' invisible window and the run hangs holding an EXCEL.EXE that blocks the next
' automation. Returning a string rather than raising also keeps a Power Query
' failure distinguishable from Excel's transient COM locks, which the Python
' caller retries.
'
' Performance toggles disable screen updates, automatic calculation, events,
' and alerts during the refresh; the original Application state is restored
' in the Cleanup block whether the refresh succeeded or raised an error.
'
' Cleanup snapshots Err BEFORE restoring that state, because a restore that
' itself fails would otherwise overwrite the real reason, and the restores are
' wrapped so one failing cannot raise out of the Function.
'
' Side effect: each WorkbookConnection's BackgroundQuery flag is set to
' False and persists in the saved workbook. This is intentional.
Function refresh() As String
    Dim conn As WorkbookConnection
    Dim prevCalc As Long
    Dim connName As String
    Dim refreshed As Long
    Dim errNum As Long
    Dim errDesc As String

    prevCalc = Application.Calculation
    connName = "(workbook setup)"

    On Error GoTo Cleanup

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For Each conn In ThisWorkbook.Connections
        connName = conn.Name
        On Error Resume Next
        conn.OLEDBConnection.BackgroundQuery = False
        ' A non-OLEDB connection leaves Err set here; without the Clear, Cleanup
        ' would report a failure even though every refresh went on to succeed
        Err.Clear
        On Error GoTo Cleanup
        conn.refresh
        refreshed = refreshed + 1
    Next conn

Cleanup:
    errNum = Err.Number
    errDesc = Err.Description
    On Error GoTo -1  ' end the active handler first, or the Resume Next below is ignored on the error path

    On Error Resume Next
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    On Error GoTo 0

    If errNum <> 0 Then
        refresh = "Power Query refresh failed on '" & connName & "' - error " & errNum & ": " & errDesc
    ElseIf refreshed = 0 Then
        refresh = "Power Query refresh did nothing - the workbook reports no connections"
    End If
End Function
