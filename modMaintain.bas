Attribute VB_Name = "modMaintain"
Option Explicit

'******************************************************************************
' Checks to see if the mrg folder is there, if not it creates it...
'
' The GazMan - November 2002
'******************************************************************************
Public Function CreateDatafolder(ByVal sFolder As String)
Dim FSO        As Scripting.FileSystemObject

On Error GoTo Err_Handler

Set FSO = New Scripting.FileSystemObject

' If FILE folder doesn't exist, create it
If Not FSO.FolderExists(App.Path & "\Files") Then
    FSO.CreateFolder App.Path & "\Files"
End If

' If folder doesn't exist, create it
If Not FSO.FolderExists(App.Path & "\Files\" & sFolder) Then
    FSO.CreateFolder App.Path & "\Files\" & sFolder
End If

Exit Function
Err_Handler:
    
    bRunFail = True
    MsgBox "An error has occurred in Function CreateDatafolder! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

Function LogEvent(strEvent)

frmRunner.lblStatus.Caption = strEvent
frmRunner.lvwLog.ListItems.Add 1, , Format(Now, "dd/mm/yyyy hh:nn:ss")
frmRunner.lvwLog.ListItems.Item(1).SubItems(1) = strEvent
If frmRunner.lvwLog.ListItems.Count = 200 Then frmRunner.lvwLog.ListItems.Remove (200)

End Function

'**********************************************************************************************
' Connect to the Access database...
'
'**********************************************************************************************
Function ConnectionStrings()
   
Dim i       As Integer

On Error GoTo Err_Handler

' Set up the Connection to the Access database....
AccessApp = Trim(App.Path & "\Runner.mdb")

Set cnList = New ADODB.Connection
cnList.Open "PROVIDER=MSDASQL;" & _
        "DRIVER={Microsoft Access Driver (*.mdb)};" & _
              "DBQ= " & AccessApp & ";" & _
              "UID=sa;PWD=;"

LogEvent "Database connections setup...."

Exit Function
Err_Handler:
    
    bRunFail = True
    MsgBox "An error has occurred in Function ConnectionStrings! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function
