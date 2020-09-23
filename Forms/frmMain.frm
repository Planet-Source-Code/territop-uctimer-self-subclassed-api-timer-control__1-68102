VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ucTimer - v1.0.0"
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5055
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3195
   ScaleWidth      =   5055
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin Project1.ucTimer ucTimer1 
      Left            =   120
      Top             =   2640
      _extentx        =   661
      _extenty        =   661
   End
   Begin VB.Frame fmProperties 
      Caption         =   "Properties:"
      Height          =   2415
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   4815
      Begin VB.OptionButton opSeries 
         Caption         =   "Count Down"
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   12
         Top             =   1800
         Width           =   1335
      End
      Begin VB.OptionButton opSeries 
         Caption         =   "Count Up"
         Height          =   255
         Index           =   0
         Left            =   1320
         TabIndex        =   11
         Top             =   1800
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.ComboBox cmbThread 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1080
         Width           =   1095
      End
      Begin VB.ComboBox cmbDuration 
         Height          =   315
         Left            =   1320
         TabIndex        =   6
         Text            =   "cmbDuration"
         Top             =   480
         Width           =   1095
      End
      Begin VB.ComboBox cmbInterval 
         Height          =   315
         Left            =   3480
         TabIndex        =   4
         Text            =   "cmbInterval"
         Top             =   480
         Width           =   1095
      End
      Begin VB.CheckBox chkEnabled 
         Caption         =   "Enabled"
         Height          =   255
         Left            =   2760
         TabIndex        =   3
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label lblTimerType 
         Caption         =   "Timer Type:"
         Height          =   255
         Left            =   360
         TabIndex        =   10
         Top             =   1800
         Width           =   855
      End
      Begin VB.Label lblThread 
         Caption         =   "Thread Priority:"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label lblDuration 
         Caption         =   "Duration:"
         Height          =   255
         Left            =   480
         TabIndex        =   7
         Top             =   480
         Width           =   735
      End
      Begin VB.Label lblInterval 
         Caption         =   "Interval:"
         Height          =   255
         Left            =   2760
         TabIndex        =   5
         Top             =   480
         Width           =   615
      End
   End
   Begin VB.Label lblResult 
      Height          =   255
      Left            =   2280
      TabIndex        =   1
      Top             =   2760
      Width           =   1335
   End
   Begin VB.Label lblLabel 
      Caption         =   "Timer Count:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1080
      TabIndex        =   0
      Top             =   2760
      Width           =   1095
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub chkEnabled_Click()
    With Me
        .ucTimer1.Enabled = (.chkEnabled.Value = vbChecked)
    End With
End Sub

Private Sub cmbDuration_Change()
    With Me
        If IsNumeric(.cmbDuration.Text) Then
            .ucTimer1.Duration = CLng(.cmbDuration.Text)
        Else
            MsgBox "Numeric Values Only!", vbInformation + vbOKOnly, "ucTimer"
            .cmbDuration.SelStart = 0
            .cmbDuration.SelLength = LenB(.cmbDuration.Text)
        End If
    End With
End Sub

Private Sub cmbDuration_Click()
    With Me
        .ucTimer1.Duration = .cmbDuration.List(.cmbDuration.ListIndex)
    End With
End Sub

Private Sub cmbInterval_Change()
    With Me
        If IsNumeric(.cmbInterval.Text) Then
            .ucTimer1.Interval = CLng(.cmbInterval.Text)
        Else
            MsgBox "Numeric Values Only!", vbInformation + vbOKOnly, "ucTimer"
            .cmbInterval.SelStart = 0
            .cmbInterval.SelLength = LenB(.cmbInterval.Text)
        End If
    End With

End Sub

Private Sub cmbInterval_Click()
    With Me
        .ucTimer1.Interval = .cmbInterval.List(.cmbInterval.ListIndex)
    End With
End Sub

Private Sub cmbThread_Click()
    With Me
        Select Case .cmbThread.ListIndex
            Case 0
                .ucTimer1.ThreadPriority = utNormal
            Case 1
                .ucTimer1.ThreadPriority = utIdle
            Case 2
                .ucTimer1.ThreadPriority = utHigh
            Case 3
                .ucTimer1.ThreadPriority = utRealTime
        End Select
    End With
End Sub

Private Sub Form_Load()
    Dim i As Long
    With Me
        .Caption = "ucTimer - " & .ucTimer1.Version
        For i = 1 To 5
            .cmbDuration.AddItem 10 ^ i
            .cmbInterval.AddItem 10 ^ i
        Next
        .cmbDuration.ListIndex = 3
        .cmbInterval.ListIndex = 0
        .ucTimer1.Interval = .cmbInterval.List(.cmbInterval.ListIndex)
        .ucTimer1.Duration = .cmbDuration.List(.cmbDuration.ListIndex)
        With .cmbThread
            .AddItem "utNormal"
            .AddItem "utIdle"
            .AddItem "utHigh"
            .AddItem "utRealTime"
            .ListIndex = 3
        End With
    End With
End Sub

Private Sub opSeries_Click(Index As Integer)
    With Me
        .ucTimer1.TimerType = Index
    End With
End Sub

Private Sub ucTimer1_Elapsed(nTime As Long)
    Me.lblResult.Caption = nTime
    Debug.Print "Elapsed Time: " & nTime
End Sub

Private Sub ucTimer1_Initailized()
    Debug.Print "ucTimer has Initailized..."
End Sub

Private Sub ucTimer1_Remaining(nTime As Long)
    Me.lblResult.Caption = nTime
    Debug.Print "Remaining Time: " & nTime
End Sub

Private Sub ucTimer1_Status(ByVal sStatus As String)
    Debug.Print "ucTimer UserControl Status: " & sStatus
End Sub

Private Sub ucTimer1_Terminated()
    Debug.Print "ucTimer has Terminated..."
End Sub

Private Sub ucTimer1_Timer()
'    Debug.Print "Timer Fired"
End Sub
