Attribute VB_Name = "taskupdate"
Option Explicit
Public updateMinutes As Long
Public updateDescription As String
Public updateCanceled As Boolean
Public updateStatus
Public updateDueDate As Date

Public Sub TaskUpdate()
    Dim objItem As Outlook.TaskItem
    Dim objApp As Outlook.Application
    Set objApp = Outlook.Application
    Set objItem = objApp.ActiveInspector.currentItem
    
    updateStatus = objItem.Status
    updateDueDate = objItem.DueDate
        
    updatetaskform.Show
    
    'Exit if the userform cancel button
    If updateCanceled = True Then Exit Sub
    
    With objItem
        .Body = "UPDATE: " & updateDescription & " - " & Now & " (" & updateMinutes & " Minutes | " & Convert_Status(updateStatus) & ")" & vbNewLine & vbNewLine & .Body
        .ActualWork = .ActualWork + updateMinutes
        .TotalWork = .TotalWork + updateMinutes
        .Status = updateStatus
        .DueDate = updateDueDate
        .Save

    End With
    
    'Clear public variables
    updateMinutes = 0
    updateDescription = vbNullString
    updateCanceled = False
    Debug.Print (updateMinutes & " | " & updateCanceled)
    
    End Sub
Sub RunTaskUpdate()
    'For some reason when I add TaskUpdate to the ribbon of my task it doesn't run. _
    It does however if I add a different sub that calls TaskUpdate.
    
    TaskUpdate
End Sub
