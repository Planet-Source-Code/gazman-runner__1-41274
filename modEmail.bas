Attribute VB_Name = "modEmail"
Option Explicit

'**************************************************************************************
' Uses Outlook to send emails notifying of any errors...
' This could be used (or variation of) to send the email...
'
' The GazMan 2002
'**************************************************************************************

Public Function eMailNotify(sError As String, eMail As String, sEvent As String)

Dim olApp           As New Outlook.Application
Dim ns              As Outlook.NameSpace
Dim myitem          As Outlook.MailItem
Dim myRecipient     As Outlook.Recipient
Dim myRecipients    As Outlook.Recipients
Dim sEmail          As String
Dim iLength         As Integer
Dim iEmail          As Integer

On Error GoTo Err_Handler

If eMail = "" Then
    MsgBox "No email address has been found!", vbExclamation, "Email Address"
    Exit Function
End If

Set ns = olApp.GetNamespace("MAPI")
Set myitem = olApp.CreateItem(olMailItem)
Set myRecipient = myitem.Recipients.Add(eMail)
Set myRecipients = myitem.Recipients
If Not myitem.Recipients.ResolveAll Then 'Check each email address...
    eMail = "default@default.com"
End If

'Build the body of the message...
If bRunFail = False Then
    myitem.Body = "The following event has run correctly: " & sEvent & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "" & sError & "" & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "The job was run at: " & Now() & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "The job type was: '" & sLetterName & "'." & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "There were " & iOpCount & " records." & vbCrLf & vbCrLf
Else
    myitem.Body = "The following event has not run correctly: " & sEvent & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "Due to error: " & sError & " " & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "The error occurred at: " & Now() & vbCrLf & vbCrLf
    myitem.Body = myitem.Body & "The job was: " & sLetterName & vbCrLf & vbCrLf
End If

myitem.Send
Set myitem = Nothing
Exit Function

Err_Handler:

    MsgBox "An error has occurred in Function eMailNotify! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function
