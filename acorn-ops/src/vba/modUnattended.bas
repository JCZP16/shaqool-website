Attribute VB_Name = "modUnattended"
'==============================================================================
' modUnattended - the scheduled run.
'
' Called from Windows Task Scheduler via 99_Scripts\Invoke-AcornNightly.ps1. It
' does only the things that are safe to do with nobody watching: refresh the
' caches, rebuild the alert list, and reconcile the controlled documents. It
' does not import, does not write to the data workbooks, does not raise jobs
' and does not touch Outlook.
'
' That restraint is the point. Anything that creates a record or moves a file
' should be done by a person who can see what it did, so that when a figure
' looks wrong on Monday there is a name against the decision rather than a
' scheduled task nobody remembers setting up.
'
' Every message that would normally be a dialog is written to the day's log
' instead, so the task always finishes rather than waiting on a click.
'==============================================================================
Option Explicit

Public Sub RunNightly()
    Dim started As Date
    started = Now
    gSilent = True

    On Error Resume Next
    LogMessage "NIGHTLY", "started"

    modSync.SyncAll
    If Err.number <> 0 Then
        LogMessage "NIGHTLY", "sync failed: " & Err.description
        Err.Clear
    End If

    modAlerts.BuildAlerts
    If Err.number <> 0 Then
        LogMessage "NIGHTLY", "alerts failed: " & Err.description
        Err.Clear
    End If

    modAdmin.DocumentControlCheck
    If Err.number <> 0 Then
        LogMessage "NIGHTLY", "document control check failed: " & Err.description
        Err.Clear
    End If

    ThisWorkbook.save
    If Err.number <> 0 Then
        LogMessage "NIGHTLY", "could not save the Console: " & Err.description
        Err.Clear
    End If

    LogMessage "NIGHTLY", "finished in " & Format$((Now - started) * 86400, "0") & "s"
    On Error GoTo 0
    gSilent = False
End Sub

' Same thing, but leaves the dialogs on - for testing the scheduled run by hand
' before trusting it to Task Scheduler.
Public Sub RunNightlyInteractive()
    gSilent = False
    modSync.SyncAll
    modAlerts.BuildAlerts
    modAdmin.DocumentControlCheck
End Sub
