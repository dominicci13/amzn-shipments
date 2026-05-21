Option Explicit

' --- refresh ---------------------------------------------------------------
' Refreshes each connection synchronously. Returns only once every query has
' finished, so the calling Python code does not need a `time.sleep()` to wait
' for background queries.
'
' Performance toggles disable screen updates, automatic calculation, events,
' and alerts during the refresh; the original Application state is restored
' in the Cleanup block whether the Sub succeeded or raised an error.
'
' Side effect: each WorkbookConnection's BackgroundQuery flag is set to
' False and persists in the saved workbook. This is intentional.
Sub refresh()
    Dim conn As WorkbookConnection
    Dim prevCalc As Long

    On Error GoTo Cleanup

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    For Each conn In ThisWorkbook.Connections
        On Error Resume Next
        conn.OLEDBConnection.BackgroundQuery = False
        On Error GoTo Cleanup
        conn.refresh
    Next conn

Cleanup:
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.Calculation = prevCalc
    Application.ScreenUpdating = True
    If Err.Number <> 0 Then
        MsgBox "refresh error " & Err.Number & ": " & Err.Description, vbExclamation
    End If
End Sub
