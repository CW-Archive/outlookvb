Attribute VB_Name = "newtask"
Option Explicit
Public isMemo As Variant
Public isPersonal As Variant
Public isPlanUpdate As Variant
Public isRFI As Variant
Public isSubmittal As Variant
Public isPricing As Variant
Public isCloseout As Variant
Public objStartDate As Date
Public objDueDate As Date
Public objSubject As String
Public objBody As String
Public canceled As Boolean

Sub NewTaskWithForm()
    'New task button with userform
    Dim objTask As Outlook.TaskItem
    Dim objMail As Outlook.MailItem
    Dim listCategories As String
    Dim objNamespace, objDestFolder

    'Dim catNamesArray As Variant

    'catNamesArray = Array("Architect PR / Memo", "Closeout", "Personal / Pet Projects", "Plan Update", "Pricing", "RFI", "Submittal")
        
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
        
        isRFI = IIf(isRFI > 0, True, False)
        isSubmittal = IIf(isSubmittal, True, False)
        isPricing = IIf(isPricing > 0, True, False)
        isCloseout = IIf(isCloseout > 0, True, False)
        

        objStartDate = objMail.ReceivedTime
        objSubject = objMail.Subject
        objBody = objMail.Body
        
        'Show and init form
        newtaskform.Show
        
        'Exit if the userform cancel button
        If canceled = True Then Exit Sub
              
        If isRFI = True Then listCategories = ", RFI"
        If isSubmittal = True Then listCategories = listCategories & ", Submittal"
        If isMemo = True Then listCategories = listCategories & ", Architect PR / Memo"
        If isPersonal = True Then listCategories = listCategories & ", Personal / Pet Projects"
        If isPlanUpdate = True Then listCategories = listCategories & ", Plan Update"
        If isPricing = True Then listCategories = listCategories & ", Pricing"
        If isCloseout = True Then listCategories = listCategories & ", Closeout"
        If Len(listCategories) > 0 Then listCategories = Right(listCategories, Len(listCategories) - 2)
        Debug.Print listCategories
        
        'Create Task
        With objTask
            .Subject = objSubject
            .StartDate = objStartDate
            .DueDate = objDueDate
            .Attachments.Add objMail
            .Body = objBody
            .Categories = listCategories
            .Save
        End With
        
        'Move Email to Archive folder
        objMail.Move objDestFolder
        
    Next
    
    Set objTask = Nothing
    Set objMail = Nothing
    
End Sub
