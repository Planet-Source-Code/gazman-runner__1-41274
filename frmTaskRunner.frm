VERSION 5.00
Begin VB.Form frmRunnerSettings 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runner Settings"
   ClientHeight    =   8940
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmTaskRunner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8940
   ScaleWidth      =   7365
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00800000&
      Height          =   8655
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.Frame Frame2 
         BackColor       =   &H00E0E0E0&
         Height          =   6975
         Left            =   120
         TabIndex        =   14
         Top             =   1560
         Width           =   6855
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   120
            TabIndex        =   24
            Top             =   6000
            Width           =   6615
            Begin VB.CommandButton cmdUpdate 
               Caption         =   "Update"
               Height          =   495
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   25
               Top             =   240
               Width           =   975
            End
            Begin VB.CommandButton cmdClose 
               Caption         =   "Close"
               Height          =   495
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   11
               Top             =   220
               Width           =   975
            End
         End
         Begin VB.TextBox txbMessage 
            Appearance      =   0  'Flat
            Height          =   885
            Left            =   2160
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            Top             =   1320
            Width           =   4455
         End
         Begin VB.ComboBox cboRefresh 
            Height          =   315
            ItemData        =   "frmTaskRunner.frx":0442
            Left            =   2160
            List            =   "frmTaskRunner.frx":0444
            TabIndex        =   9
            Text            =   "24:00"
            Top             =   3720
            Width           =   975
         End
         Begin VB.CheckBox chkTest 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Test"
            Height          =   255
            Left            =   1605
            TabIndex        =   10
            Top             =   4200
            Value           =   1  'Checked
            Width           =   735
         End
         Begin VB.TextBox txbTime 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   8
            Text            =   "5"
            Top             =   3360
            Width           =   495
         End
         Begin VB.TextBox txbTo 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   2
            Text            =   "me@whatever.com"
            Top             =   600
            Width           =   2535
         End
         Begin VB.TextBox txbSubject 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   3
            Text            =   "Runner Subject"
            Top             =   960
            Width           =   2535
         End
         Begin VB.TextBox txbName 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   1
            Text            =   "The GazMan"
            Top             =   240
            Width           =   2535
         End
         Begin VB.TextBox txbProfile 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   6
            Text            =   "PROFILE"
            Top             =   2640
            Width           =   2535
         End
         Begin VB.TextBox txbServer 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   5
            Text            =   "YOURSERVER"
            Top             =   2280
            Width           =   2535
         End
         Begin VB.TextBox txbPassword 
            Appearance      =   0  'Flat
            Height          =   285
            Left            =   2160
            TabIndex        =   7
            Text            =   "PASSWORD"
            Top             =   3000
            Width           =   2535
         End
         Begin VB.Label Label3 
            BackColor       =   &H80000009&
            BackStyle       =   0  'Transparent
            Caption         =   "Copyright The Gazman 2002"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   4920
            TabIndex        =   26
            Top             =   5760
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Message"
            Height          =   255
            Index           =   13
            Left            =   960
            TabIndex        =   23
            Top             =   1320
            Width           =   975
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Document Refresh Hr"
            Height          =   255
            Index           =   12
            Left            =   120
            TabIndex        =   22
            Top             =   3720
            Width           =   1815
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Check every (mins)"
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   21
            Top             =   3360
            Width           =   1695
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "To"
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   20
            Top             =   600
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Subject"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   19
            Top             =   960
            Width           =   615
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Name"
            Height          =   255
            Index           =   7
            Left            =   1200
            TabIndex        =   18
            Top             =   240
            Width           =   735
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Exchange Server"
            Height          =   255
            Index           =   4
            Left            =   600
            TabIndex        =   17
            Top             =   2280
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Password"
            Height          =   255
            Index           =   5
            Left            =   600
            TabIndex        =   16
            Top             =   3000
            Width           =   1335
         End
         Begin VB.Label Label1 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            Caption         =   "Profile"
            Height          =   255
            Index           =   6
            Left            =   600
            TabIndex        =   15
            Top             =   2640
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Settings"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   13
         Top             =   840
         Width           =   6855
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "Runner"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   24
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   615
         Left            =   120
         TabIndex        =   12
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   6720
      Picture         =   "frmTaskRunner.frx":0446
      Top             =   3240
      Width           =   510
   End
End
Attribute VB_Name = "frmRunnerSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()

    Visible = False 'Just make the form invisible, so we can still access the variable....

End Sub

Private Sub cmdUpdate_Click()
Dim cnSettings      As New ADODB.Connection

On Error GoTo Err_Handler

sSQL = "Update RunnerSettings "
sSQL = sSQL & "SET "
sSQL = sSQL & "TestToAddr ='" & txbTo & "', "
sSQL = sSQL & "Subject = '" & txbSubject & "', Message = '" & txbMessage & "', "
sSQL = sSQL & "Server = '" & txbServer & "', Profile = '" & txbProfile & "', "
sSQL = sSQL & "Password ='" & txbPassword & "', RefreshDocs ='" & cboRefresh.Text & "', "
sSQL = sSQL & "Test = '" & chkTest.Value & "', PollTime = '" & txbTime.Text & "'  "

cnList.Execute sSQL
MsgBox "Settings updated....", vbInformation, "Settings"

LogEvent "Settings updated...."

Exit Sub
Err_Handler:

    bRunFail = True
    MsgBox "An error has occurred in Sub cmdUpdate_Click! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Sub

Public Function GetSettings()

Dim rsSet    As ADODB.Recordset
Dim sTable      As String

On Error GoTo Err_Handler

sSQL = "SELECT * From RunnerSettings"
Set rsSet = New ADODB.Recordset
rsSet.Open sSQL, cnList, adOpenForwardOnly, adLockReadOnly

If rsSet.EOF = False Then
    txbMessage.Text = rsSet.Fields("Message")
    txbPassword.Text = rsSet.Fields("Password")
    txbProfile.Text = rsSet.Fields("Profile")
    txbServer.Text = rsSet.Fields("Server")
    txbSubject.Text = rsSet.Fields("Subject")
    txbTime.Text = rsSet.Fields("PollTime")
    txbTo.Text = rsSet.Fields("TestToAddr")
    cboRefresh.Text = rsSet.Fields("RefreshDocs")
    chkTest.Value = rsSet.Fields("Test")
End If

rsSet.Close
Set rsSet = Nothing

Exit Function
Err_Handler:
    
    bRunFail = True
    MsgBox "An error has occurred in Function GetSettings! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation

End Function

Private Sub Form_Load()
Dim i As Integer

i = 1
Do Until i = 25
    cboRefresh.AddItem i & ":00", i - 1
    i = i + 1
Loop

End Sub
