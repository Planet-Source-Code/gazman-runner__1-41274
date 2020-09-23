VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmRunner 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Runner"
   ClientHeight    =   8925
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7365
   Icon            =   "frmFaxRunner.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8925
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
         TabIndex        =   7
         Top             =   1560
         Width           =   6855
         Begin VB.TextBox txbIndicate1 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   195
            Left            =   6000
            TabIndex        =   14
            Top             =   5400
            Width           =   255
         End
         Begin VB.TextBox txbIndicate2 
            Appearance      =   0  'Flat
            BackColor       =   &H0080FF80&
            Height          =   195
            Left            =   6240
            TabIndex        =   13
            Top             =   5400
            Width           =   255
         End
         Begin VB.Frame Frame4 
            BackColor       =   &H00E0E0E0&
            Height          =   855
            Left            =   120
            TabIndex        =   9
            Top             =   5520
            Width           =   6615
            Begin VB.CommandButton cmdClose 
               Caption         =   "Close"
               Height          =   495
               Left            =   5400
               Style           =   1  'Graphical
               TabIndex        =   12
               Top             =   220
               Width           =   975
            End
            Begin VB.CommandButton cmdSettings 
               Caption         =   "Settings"
               Height          =   495
               Left            =   4200
               Style           =   1  'Graphical
               TabIndex        =   3
               Top             =   220
               Width           =   975
            End
            Begin VB.CommandButton cmdStart 
               Caption         =   "Start"
               Height          =   495
               Left            =   120
               Style           =   1  'Graphical
               TabIndex        =   2
               Top             =   220
               Width           =   975
            End
            Begin VB.CommandButton cmdStop 
               Caption         =   "Stop"
               Height          =   495
               Left            =   1320
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   220
               Width           =   975
            End
         End
         Begin VB.Frame Frame3 
            BackColor       =   &H00E0E0E0&
            Height          =   615
            Left            =   120
            TabIndex        =   10
            Top             =   6240
            Width           =   6615
            Begin VB.Timer Timer2 
               Left            =   5160
               Top             =   120
            End
            Begin VB.Timer Timer1 
               Left            =   5640
               Top             =   120
            End
            Begin VB.TextBox txbStatus 
               Appearance      =   0  'Flat
               BackColor       =   &H000000FF&
               Height          =   285
               Left            =   6120
               TabIndex        =   11
               Top             =   210
               Width           =   255
            End
            Begin VB.Label lblStatus 
               BackColor       =   &H00E0E0E0&
               Height          =   255
               Left            =   120
               TabIndex        =   6
               Top             =   240
               Width           =   6135
            End
         End
         Begin MSComctlLib.ListView lvwLog 
            Height          =   5055
            Left            =   120
            TabIndex        =   1
            Top             =   240
            Width           =   6615
            _ExtentX        =   11668
            _ExtentY        =   8916
            View            =   3
            LabelWrap       =   -1  'True
            HideSelection   =   -1  'True
            FullRowSelect   =   -1  'True
            _Version        =   393217
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   2
            BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               Text            =   "Date/Time"
               Object.Width           =   2999
            EndProperty
            BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
               SubItemIndex    =   1
               Text            =   "Event"
               Object.Width           =   8290
            EndProperty
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         Caption         =   "eMail Word Documents"
         BeginProperty Font 
            Name            =   "Garamond"
            Size            =   18
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
         TabIndex        =   8
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
         TabIndex        =   5
         Top             =   240
         Width           =   6855
      End
   End
   Begin VB.Image Image3 
      Height          =   510
      Left            =   6720
      Picture         =   "frmFaxRunner.frx":0442
      Top             =   3240
      Width           =   510
   End
End
Attribute VB_Name = "frmRunner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim i As Integer


Private Sub KickOff()

On Error GoTo ErrorTrap
  
Opportunity 'Check for records ready....

Exit Sub

ErrorTrap:

    MsgBox "An error has occurred in Sub KickOff! " & vbCrLf & vbCrLf & Err & ": " & Error & " ", vbExclamation
    
End Sub

Private Sub cmdClose_Click()

cmdStop_Click
Unload Me
Unload frmRunnerSettings

End Sub

Private Sub cmdSettings_Click()

frmRunnerSettings.GetSettings
Load frmRunnerSettings
frmRunnerSettings.Show vbModal, Me

End Sub

Private Sub cmdStart_Click()

If frmRunnerSettings.cboRefresh.Text = "" Then
    MsgBox "Please select a refresh time...", vbExclamation, "Refresh Time"
    frmRunnerSettings.cboRefresh.SetFocus
    Exit Sub
End If

If i = 0 Then
    MapiLogon
    If bLoggedOn = False Then Exit Sub
    Timer1.Interval = 1
Else
    Timer1.Interval = 60000
End If
Timer2.Interval = 1000

End Sub

Private Sub cmdStop_Click()

Timer1.Interval = 0: Timer2.Interval = 0
MAPILogOff
LogEvent "Exchange Server log off completed..."
frmRunner.txbStatus.BackColor = &HFF&
bLoggedOn = False
i = 0

End Sub

Private Sub Form_Load()
Dim i As Integer

ConnectionStrings
frmRunnerSettings.GetSettings

End Sub

Private Sub Timer1_Timer()

If bLoggedOn = True Then
    If bRunFail = True Then Timer1.Interval = 0
    If i = 0 Then 'If just started then allow to run...
        i = i + 1
        Timer1.Interval = 60000
        KickOff
    Else
        If i = CInt(frmRunnerSettings.txbTime.Text) Then 'Wait X minutes until run again...
            i = 0
            KickOff
            LogEvent "Checking for records again in " & CInt(frmRunnerSettings.txbTime.Text) & " minute(s)."
        Else
            If frmRunnerSettings.txbTime.Text = 1 Then
                LogEvent "Checking for faxes again in " & CInt(frmRunnerSettings.txbTime.Text) & " minute."
            Else
                LogEvent "Checking for records again in " & CInt(frmRunnerSettings.txbTime.Text) - i & " minute(s)."
            End If
            i = i + 1
        End If
    End If
End If

End Sub

Private Sub Timer2_Timer()

If txbIndicate1.BackColor = 8454016 Then
    txbIndicate1.BackColor = 16777215
    txbIndicate2.BackColor = 8454016
Else
    txbIndicate1.BackColor = 8454016
    txbIndicate2.BackColor = 16777215
End If

End Sub
