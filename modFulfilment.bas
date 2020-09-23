Attribute VB_Name = "modFulfilment"
Option Explicit

Public Declare Function GetPrivateProfileString Lib "Kernel32" _
    Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, lpKeyName _
    As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, _
    ByVal nSize As Integer, ByVal lpFileName As String) As Integer

Public wdApp                    As Word.Application

Public cnList                   As ADODB.Connection

Public sFieldNames              As String
Public sTableNames              As String
Public sLetterName              As String

Public sSQL                     As String
Public sSQLDTS                  As String
Public sCustID                  As String
Public sCustIDUpdate            As String
Public Profile                  As String
Public MailServer               As String
Public sEmail                   As String
Public AccessApp                As String
Public sLocalFile               As String

Public iTable                   As Integer
Public iIndex                   As Integer
Public iOpCount                 As Integer
Public iBranch                  As Integer
Public iFax                     As Integer
Public iWordCount               As Integer

Public bRunFail                 As Boolean
Public bLoggedOn                As Boolean

Public vFields                  As Variant

'************************************************************************
' Run the main functions from here...
'
'************************************************************************
Public Function CreateFile()
Dim iLine   As Integer

On Error GoTo Err_Handler

'Build the SQL required to supply all the data for the fulfilment files....
sSQLDTS = BuildSQL

LogEvent "SQL Query built...."
If bRunFail = True Then Exit Function

'Create .MRG file....this has all the data required for the mergefields...Address etc.
CreateMRGfile

Exit Function
Err_Handler:

    bRunFail = True
    MsgBox "An error has occurred in Function CreateFile! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'*****************************************************************************************
' This function creates a comma delimited text file (.mrg) that contains data required
' by the fulfilment system to merge with the documents. Addr1, Fname etc
' With the rs from sSQLDTS, create a comma delimited text file for the data.
'
' The GazMan 2002
'*****************************************************************************************
Function CreateMRGfile()

Dim rsData          As ADODB.Recordset

Dim sHeader         As String
Dim sOutPutData     As String

Dim sError          As String
Dim sFilePath       As String
Dim sCallDte        As String

Dim iDataFile       As Integer
Dim nx              As Integer
Dim ny              As Integer
Dim nFnd            As Integer
Dim nColumns        As Integer
Dim nRows           As Integer
Dim iPos            As Integer
Dim iLength         As Integer

Dim vRsGetRow()     As Variant

On Error GoTo Err_Handler

Set rsData = New ADODB.Recordset
rsData.Open sSQLDTS, cnList, adOpenStatic, adLockReadOnly

If rsData.RecordCount = 0 Then GoTo Err_Handler:
sLocalFile = App.Path & "\Files\mrg\" & sLetterName & ".mrg"

vRsGetRow = rsData.GetRows
nColumns = UBound(vRsGetRow(), 1) 'Gets the headers
nRows = UBound(vRsGetRow(), 2) 'The data for the headers
iDataFile = FreeFile

'Check that the local directory exists and then write a copy of the text file.
CreateDatafolder "mrg"
Open sLocalFile For Output As iDataFile

'Create the column headers...
For ny = 0 To nColumns
    'Headers added here...
    sHeader = Chr(34) & rsData.Fields.Item(ny).Name & Chr(34)
    If ny = nColumns Then
        Print #iDataFile, sHeader
    Else
        Print #iDataFile, sHeader & Chr(44);
    End If
Next ny

'Add the merge fields data....
For nx = 0 To nRows
    
    For ny = 0 To nColumns 'Not the last field...
        If IsNull(vRsGetRow(ny, nx)) Then vRsGetRow(ny, nx) = ""
        If ny = nColumns Then 'The last field, indicate new record line...
            If IsNull(vRsGetRow(ny, nx)) = True Then
                Print #iDataFile, Chr(34) & "" & Chr(34)
            Else
                Print #iDataFile, Chr(34) & CStr(vRsGetRow(ny, nx)) & Chr(34)
            End If
            GoTo NextNY
            
        Else '*******************************************************THE GREAT DIVIDE
        
            If IsNull(vRsGetRow(ny, nx)) = True Then
                Print #iDataFile, Chr(34) & "" & Chr(34) & Chr(44);
            Else
                Print #iDataFile, Chr(34) & CStr(vRsGetRow(ny, nx)) & Chr(34) & Chr(44);
            End If
            GoTo NextNY
        End If
NextNY:
    Next ny
Next nx
Close iDataFile

rsData.Close
Set rsData = Nothing

LogEvent "Merge file created....."

Exit Function
Err_Handler:
    
    bRunFail = True
    MsgBox "An error has occurred in Function CreateMRGfile! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'*****************************************************************************************
' Updates the status in the Fax table
'
' The GazMan 2002
'*****************************************************************************************
Function UpdateTables(sOptyIDUpdate As String)

Dim cnUpdateOppTable    As ADODB.Connection
Dim sSQL                As String
Dim bFul                As Boolean
Dim bOpp                As Boolean
Dim bFax                As Boolean

On Error GoTo Err_Handler

'Update the Runner table for the OptyID
sSQL = "Update Runner "
sSQL = sSQL & "SET CompleteDTE = '" & Now() & "', Status = 'SENT' "
sSQL = sSQL & "Where (CustID = " & sCustIDUpdate & ")"

cnList.Execute sSQL
bFax = True

sCustID = ""
sCustIDUpdate = ""

Exit Function
Err_Handler:

        bRunFail = True
        MsgBox "An error has occurred in Function UpdateTables!" & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    
End Function

