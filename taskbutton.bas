Attribute VB_Name = "taskbutton"
Option Explicit
Sub MailtoTask()
    'Written by Camron Walker
    
    Dim objTask As Outlook.TaskItem
    Dim objMail As Outlook.MailItem
    Dim objNamespace, objDestFolder, objTaskId
    Dim isRFI As Integer, isSubmittal As Integer, isPricing As Integer, isCloseout As Integer, dueDays As Integer
    Dim listCategories As String, objSubject As String, inputAnswer As String
    
    'Check to make sure there is an email selected
    If Application.ActiveExplorer.Selection.count = 0 Then
        MsgBox ("No item selected")
        Exit Sub
    End If
    
    'Set Arcive as folder to move email to
    Set objNamespace = Application.GetNamespace("MAPI")
    Set objDestFolder = objNamespace.Folders("camron@gowestland.com").Folders("Archive")
    
    For Each objMail In Application.ActiveExplorer.Selection
        Set objTask = Application.CreateItem(olTaskItem)
        
        'Automanically add categories TODO: This could probably be slimmed down code wise.
        isRFI = InStr(1, objMail.Subject, "RFI", 1) + InStr(1, objMail.Body, "RFI", 1)
        isSubmittal = InStr(1, objMail.Subject, "Submittal", 1) + InStr(1, objMail.Body, "Submittal", 1)
        isPricing = InStr(1, objMail.Subject, "Pricing", 1) + InStr(1, objMail.Subject, "Quote", 1) + InStr(1, objMail.Body, "Pricing", 1) + InStr(1, objMail.Body, "Quote", 1)
        isCloseout = InStr(1, objMail.Subject, "Closeout", 1) + InStr(1, objMail.Body, "Closeout", 1) + InStr(1, objMail.Subject, "Warranty", 1) + InStr(1, objMail.Body, "Warranty", 1)
        
        If isRFI > 0 Then listCategories = ", RFI"
        If isSubmittal > 0 Then listCategories = listCategories & ", Submittal"
        If isPricing > 0 Then listCategories = listCategories & ", Pricing"
        If isCloseout > 0 Then listCategories = listCategories & ", Closeout"
        If Len(listCategories) > 0 Then listCategories = Right(listCategories, Len(listCategories) - 2)
        Debug.Print listCategories
        
        'Check subject and duedate TODO: Think of some way to skip this if I don't want to do it.
        dueDays = 7
        objSubject = objMail.Subject
        
        inputAnswer = InputBox("Edit the duedate or TaskName. (DaysUntilDue - Subject name)", "Create Task", dueDays & "-&&-" & objSubject)
        
        objSubject = Right(inputAnswer, Len(inputAnswer) - InStr(1, inputAnswer, "-&&-", 1) - 3)
        dueDays = Left(inputAnswer, InStr(1, inputAnswer, "-&&-", 1) - 1)
               
        'Create Task
        With objTask
            .Subject = objSubject
            .StartDate = objMail.ReceivedTime
            .DueDate = Now + dueDays
            .Attachments.Add objMail
            .Body = objMail.Body
            .Categories = listCategories
            .Save
        End With
        
        'Move Email to Archive folder
        objMail.Move objDestFolder
    Next
    
    Set objTask = Nothing
    Set objMail = Nothing

End Sub
