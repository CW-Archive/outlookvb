Attribute VB_Name = "tasktime"
Option Explicit
Sub AddOneHourActivity()
    AddTimeActivity (60)
End Sub

Sub AddHalfHourActivity()
    AddTimeActivity (30)
End Sub

Sub AddQuarterHourActivity()
    AddTimeActivity (15)
End Sub
Sub AddOtherTimeActivity()
    Dim inputResult As Long
    inputResult = InputBox("How much time? (In minutes).", "Time Amount", 120)
    AddTimeActivity (inputResult)
End Sub

Sub AddTimeActivity(timeAmount As Long)
    Dim objTask As Outlook.TaskItem
    Dim workDiscription As String
        
    'Check to make sure there is an email selected
    If Application.ActiveExplorer.Selection.count = 0 Then
        MsgBox ("No item selected")
        Exit Sub
    End If

    For Each objTask In Application.ActiveExplorer.Selection
        With objTask
            workDiscription = InputBox("What did you do during this time? (Can't be blank).", "Track Time")
            If workDiscription = vbNullString Then GoTo nextWith
        
            .Body = workDiscription & " - " & Now & vbNewLine & vbNewLine & .Body
            .ActualWork = .ActualWork + timeAmount
            .TotalWork = .TotalWork + timeAmount
            .Save
nextWith:
        End With
    Next
End Sub
