Attribute VB_Name = "newtask"
'Option Explicit
Public isRFI As Variant
Public isSubmittal As Variant
Public isPricing As Variant
Public isCloseout As Variant
Public objStartDate As Date
Public objDueDate As Date
Public objSubject As String
Public objBody As String
Public listCategories As String

Sub NewTaskWithForm()
    'New task button with userform
    Dim objTask As Outlook.TaskItem
    Dim objMail As Outlook.MailItem

    Dim catNamesArray As Variant, listVarArray

    
    catNamesArray = Array("Architect PR / Memo", "Closeout", "Personal / Pet Projects", "Plan Update", "Pricing", "RFI", "Submittal")
        
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
