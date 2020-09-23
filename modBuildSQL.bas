Attribute VB_Name = "modBuildSQL"
Option Explicit

Dim sSQLa           As String
Dim aTables(10)     As String 'Array of tables for the

'********************************************************************************************
' Take the table and get the indexes and primary key. Need to Join with the appropriate
' index/primary key on the next table and so on. ie; build a relevant set of joins between
' tables that have been selected.                                                                                                                                           *
'
' The GazMan May 2002
'********************************************************************************************

Public Function BuildSQL()
Dim rsBuildSQL      As ADODB.Recordset
Dim rsSelClause     As ADODB.Recordset
Dim sTableAlias     As String
Dim sTableNames     As String
Dim sSelClause      As String
Dim sTableNameNext  As String
Dim iRecords        As Integer
Dim iTable          As Integer
Dim i               As Integer
Dim vTable          As Variant
Dim bTables         As Boolean

On Error GoTo Err_Handler

sSQL = "SELECT Document, Table, Field "
sSQL = sSQL & "From TableFields "
sSQL = sSQL & "WHERE (((TableFields.Document)= '" & sLetterName & "' ));"

Set rsBuildSQL = New ADODB.Recordset
rsBuildSQL.Open sSQL, cnList, adOpenKeyset, adLockReadOnly

If rsBuildSQL.RecordCount = 0 Then GoTo Err_Handler

With rsBuildSQL
    Do Until .EOF 'Get all the fields and tables required....
        If iRecords = 0 Then
            sFieldNames = .Fields("Table") & "." & .Fields("Field")
            vTable = .Fields("Table")
            aTables(iTable) = vTable 'Add table name to table array...one possibility
            sTableNames = .Fields("Table") 'here's the other...
            iTable = iTable + 1
        Else
            sTableNameNext = .Fields("Table")
            sFieldNames = sFieldNames & ", " & sTableNameNext & "." & .Fields("Field")
        
            If InStr(1, sTableNames, sTableNameNext, vbTextCompare) = False Then 'New table?
                vTable = sTableNameNext
                aTables(iTable) = vTable
                sTableNames = sTableNameNext & ", " & sTableNames
                iTable = iTable + 1
                bTables = True
            End If
        End If
        iRecords = iRecords + 1
        .MoveNext
    Loop
    .Close
End With

BuildSQLaCase 'Add the necessary Join(s) to the statement

' OK, now we have the fields, tabel joins etc, put it all together...
sSQL = "SELECT DISTINCT "
sSQL = sSQL & "" & sFieldNames & " "
sSQL = sSQL & "FROM Customer "
sSQL = sSQL & sSQLa
sSQL = sSQL & "WHERE (" & sCustID & " ) "

BuildSQL = sSQL
sSQLa = ""

Exit Function
Err_Handler:
    
    bRunFail = True
    MsgBox "An error has occurred in Function BuildSQL! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    
End Function

'**********************************************************************************
' Standard set of JOINS for the relevant tables...
'
' The GazMan - May 2002
'**********************************************************************************

Public Function BuildSQLaCase()

Dim vData       As Variant
Dim isOkay      As Boolean

On Error GoTo Err_Handler

For Each vData In aTables
    If vData = "" Then Exit For
    Select Case vData
            
        Case "Customer"
        
            'Is the 'Lead' table in the SQL, so no need to Join
            
        Case "CustomerData"
        
            sSQLa = sSQLa & "INNER JOIN " & vData & " ON "
            sSQLa = sSQLa & "Customer.CustID = " & vData & ".CustID "
        
        'Case All the other tables you might have....etc
                
    End Select
Next

Erase aTables 'Clear the table array

Exit Function
Err_Handler:

    bRunFail = True
    MsgBox "An error has occurred in Function BuildSQLaCase! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

'******************************************************************************
' Check the Runner table to see if there is data ready..
'
' The GazMan May 2002
'******************************************************************************
Public Function Opportunity() As String
Dim rsOppty         As ADODB.Recordset
Dim rsEvent         As ADODB.Recordset
Dim sEventcode      As String
Dim sErrorMessage   As String
Dim i               As Integer

On Error GoTo Err_Handler

' Select from the table that holds the records...
sSQL = "SELECT * "
sSQL = sSQL & "From Runner "
sSQL = sSQL & "WHERE (Status = 'READY') "

Set rsEvent = New ADODB.Recordset
rsEvent.Open sSQL, cnList, adOpenKeyset, adLockReadOnly

'Tool through each record in 'READY' status..
If rsEvent.EOF = False Then
    Do Until rsEvent.EOF = True
            sLetterName = rsEvent.Fields("JobType")
            'Find the record in the Customer
            sSQL = "SELECT * "
            sSQL = sSQL & "From Customer "
            sSQL = sSQL & "WHERE (CustID = " & rsEvent.Fields("CustID") & ") " 'And (Status = 'READY')
                        
            Set rsOppty = New ADODB.Recordset
            rsOppty.Open sSQL, cnList, adOpenKeyset, adLockReadOnly
            
                iOpCount = rsOppty.RecordCount
                If rsOppty.EOF = False Then
                     
                ' A bit stodgy this, but works....
                sCustID = "Customer.CustID = " & rsEvent.Fields("CustID") & ""
                sCustIDUpdate = rsEvent.Fields("CustID")
                
                LogEvent "Processing Customer ID " & sCustID & ""
            End If
            
            If sCustID = "" Or bRunFail = True Then GoTo NextRecord:
            
            CreateFile 'Create the data file required....
                   
            SendMail     'Get that puppy out the door....
                   
            If bRunFail = True Then GoTo NextRecord
            UpdateTables sCustIDUpdate  'Update tables
NextRecord:
        bRunFail = False
        rsEvent.MoveNext
    Loop
Else
    LogEvent "There are no records to process!"
End If

rsEvent.Close
Set rsEvent = Nothing

Exit Function
Err_Handler:
    
    bRunFail = True
    MsgBox "An error has occurred in Function Opportunity! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

