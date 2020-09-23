Attribute VB_Name = "modSendFax"
Option Explicit

'MAPI objects.....
Dim objSession          As MAPI.Session
Dim objMessages         As MAPI.Messages
Dim objOneMessage       As MAPI.Message

Dim strEmailTemp        As String

Private boolUseCurrentSession, boolLogonDialog

'******************************************************************************
' Logon to the Exchange Server with the account specified
'
' The GazMan - November 2002
'******************************************************************************
Public Function MapiLogon()

Dim strProfileInfo      As String
Dim i                   As Integer
Dim bstrPublicRootID    As Boolean
Dim bSent               As Boolean

On Error GoTo ErrorHandler

LogEvent "Logging on to the '" & frmRunnerSettings.txbServer & "' Exchange Server..."

strProfileInfo = frmRunnerSettings.txbServer & vbLf & frmRunnerSettings.txbProfile
Set objSession = CreateObject("MAPI.Session")
objSession.Logon , , , True, , True, strProfileInfo

If (Err.Number <> 0) Or (objSession.CurrentUser.Name = "Unknown") Then
    objSession.Logoff 'Not a good logon, logoff and exit
    MsgBox "Logon error!", vbOKOnly + vbExclamation, "MapiLogon"
    LogEvent "Exchange Server Logon NOT successful...."
    frmRunner.txbStatus.BackColor = &HFF&
    frmRunner.Timer1.Interval = 0: frmRunner.Timer2.Interval = 0
    bLoggedOn = False
    bRunFail = True
    Exit Function
Else
    LogEvent "Exchange Server Logon successful...."
    frmRunner.txbStatus.BackColor = &HC000&
    bLoggedOn = True
End If

Exit Function
ErrorHandler:

    bRunFail = True
    MsgBox "An error has occurred in Function MapiLogon! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

Function MAPILogOff()

Set objSession = Nothing
Set objOneMessage = Nothing
Set objMessages = Nothing
Set objSession = Nothing

End Function

'******************************************************************************
' Sends the email....
'
' The GazMan - November 2002
'******************************************************************************
Function SendMail()

On Error GoTo ErrorHandler

Set objOneMessage = objSession.Outbox.Messages.Add
objOneMessage.Update

With objOneMessage
    .Importance = mapiHigh
    .Subject = frmRunnerSettings.txbSubject & " email for " & sCustIDUpdate
    .Text = frmRunnerSettings.txbMessage.Text
End With
With objOneMessage.Recipients.Add
    If frmRunnerSettings.chkTest.Value = 1 Then
        .Name = frmRunnerSettings.txbTo
    Else
        .Name = sEmail
    End If
    .Type = mapiTo
    .Resolve
End With
If Len(sLocalFile) Then
    MergeDoc sLocalFile 'Merge the doc to the data...
    With objOneMessage.Attachments.Add
        .Name = sLetterName
        .Type = mapiFileData
        .Source = strEmailTemp
        .ReadFromFile strEmailTemp
        .Position = 2880
    End With
End If
objOneMessage.Send

If frmRunnerSettings.chkTest = 1 Then
    frmRunner.lblStatus.Caption = "Email sent to " & frmRunnerSettings.txbTo.Text & "...."
    LogEvent "Email sent to " & frmRunnerSettings.txbTo.Text & "...."
Else
    frmRunner.lblStatus.Caption = "Email sent to " & sEmail & "...."
    LogEvent "Email sent to " & sEmail & "...."
End If

iWordCount = iWordCount + 1
If iWordCount > 6 Then 'Refresh Word after six cycles...
    wdApp.Quit
    iWordCount = 0
End If

Exit Function
ErrorHandler:

    bRunFail = True
    MsgBox "An error has occurred in Function SendMail! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'******************************************************************************
' This grabs the .mrg file created and merges the data with the
' document ready to be sent as an attachment...
'
' The GazMan - November 2002
'******************************************************************************

Function MergeDoc(sFaxData As String)

Dim strDocPath      As String
Dim strLocProjects  As String
Dim iPages          As Integer

On Error GoTo ErrorHandler

'Find the document required and make a temporary copy of it...
strLocProjects = App.Path & "\" & sLetterName & ".doc"
strDocPath = App.Path & "\Temp.doc"
strEmailTemp = App.Path & "\" & sLetterName & "_email.doc"
FileCopy strLocProjects, strDocPath

IsWordActive 'Check the status of Word...

wdApp.Documents.Open strDocPath
With wdApp.ActiveDocument.MailMerge
    .OpenDataSource sFaxData
End With

With wdApp.ActiveDocument.MailMerge
    .Destination = wdSendToNewDocument
    .Execute
    With wdApp
        .ActiveDocument.SaveAs strEmailTemp
        .ActiveDocument.Repaginate
        iPages = .ActiveDocument.BuiltInDocumentProperties(14)
        .ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges 'Close the first doc.
        .ActiveDocument.Close SaveChanges:=wdDoNotSaveChanges 'Close the second doc.
        LogEvent "Data merged to document...."
    End With
End With

Exit Function
ErrorHandler:

    bRunFail = True
    MsgBox "An error has occurred in Function MergeDoc! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'******************************************************************************
' Checks to see if Word is open, if not it cranks it up otherwise...sweeeet as
'
' The GazMan - November 2002
'******************************************************************************
Function IsWordActive()

On Error Resume Next

If wdApp.Name <> "" Then
End If

If Err.Number = 462 Or 91 Then
    Err.Clear
    With wdApp
        Set wdApp = GetObject(, "Word.Application")
        If Err.Number <> 0 Then
            Set wdApp = CreateObject("Word.Application")
        End If
    End With
End If

wdApp.Visible = True

On Error GoTo 0

End Function
