VERSION 5.00
Begin VB.Form frmQuery 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Drive Indexing Options"
   ClientHeight    =   3645
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5565
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmQuery.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3645
   ScaleWidth      =   5565
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdSet 
      Caption         =   "Start Indexing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   0
      Left            =   3870
      TabIndex        =   0
      Top             =   2910
      Width           =   1515
   End
   Begin VB.CheckBox chkIndex 
      Caption         =   "Enable Auto Indexing"
      Height          =   225
      Left            =   150
      TabIndex        =   16
      Top             =   840
      Width           =   2205
   End
   Begin VB.PictureBox picOptions 
      BorderStyle     =   0  'None
      Height          =   1215
      Left            =   150
      ScaleHeight     =   1215
      ScaleWidth      =   2505
      TabIndex        =   10
      Top             =   1140
      Width           =   2505
      Begin VB.ComboBox cbDays 
         Height          =   315
         ItemData        =   "frmQuery.frx":038A
         Left            =   840
         List            =   "frmQuery.frx":038C
         TabIndex        =   11
         Text            =   "Combo1"
         Top             =   870
         Width           =   1335
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "12 Hours"
         Height          =   225
         Index           =   0
         Left            =   0
         TabIndex        =   14
         Top             =   300
         Width           =   975
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "24 Hours"
         Height          =   225
         Index           =   1
         Left            =   0
         TabIndex        =   13
         Top             =   630
         Value           =   -1  'True
         Width           =   1065
      End
      Begin VB.OptionButton optIndex 
         Caption         =   "Every:"
         Height          =   225
         Index           =   2
         Left            =   0
         TabIndex        =   12
         Top             =   960
         Width           =   1065
      End
      Begin VB.Label lblNotice 
         AutoSize        =   -1  'True
         Caption         =   "Interval:"
         Height          =   195
         Left            =   0
         TabIndex        =   15
         Top             =   60
         Width           =   630
      End
   End
   Begin VB.CheckBox chkStart 
      Caption         =   "Run From System Tray"
      Height          =   225
      Left            =   180
      TabIndex        =   9
      Top             =   2790
      Value           =   1  'Checked
      Width           =   2055
   End
   Begin VB.CommandButton cmdSet 
      Caption         =   "Finish"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Index           =   1
      Left            =   3870
      TabIndex        =   5
      Top             =   2910
      Visible         =   0   'False
      Width           =   1515
   End
   Begin VB.DriveListBox Drive1 
      Height          =   315
      Left            =   3240
      TabIndex        =   1
      Top             =   1230
      Width           =   1965
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Status: Scanning Drive"
      Height          =   195
      Index           =   1
      Left            =   150
      TabIndex        =   17
      Top             =   3390
      Width           =   1635
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      Caption         =   "Scan Time Est:"
      Height          =   195
      Index           =   2
      Left            =   2640
      TabIndex        =   8
      Top             =   2340
      Width           =   1050
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      Caption         =   "Used Space:"
      Height          =   195
      Index           =   1
      Left            =   2790
      TabIndex        =   7
      Top             =   2070
      Width           =   900
   End
   Begin VB.Label lblStats 
      AutoSize        =   -1  'True
      Caption         =   "Drive:"
      Height          =   195
      Index           =   0
      Left            =   3240
      TabIndex        =   6
      Top             =   1800
      Width           =   435
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Start Indexing"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   165
      Index           =   4
      Left            =   1650
      TabIndex        =   4
      Top             =   510
      Width           =   1080
   End
   Begin VB.Label lblInfo 
      Caption         =   $"frmQuery.frx":038E
      Height          =   795
      Index           =   3
      Left            =   180
      TabIndex        =   3
      Top             =   90
      Width           =   5115
   End
   Begin VB.Label lblInfo 
      AutoSize        =   -1  'True
      Caption         =   "Select a Drive"
      Height          =   195
      Index           =   2
      Left            =   3240
      TabIndex        =   2
      Top             =   1020
      Width           =   990
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_lCount                As Long
Private m_lRefNum               As Long
Private m_dStatCnt              As Double
Private m_bSaveIndex            As Boolean
Private m_bChange               As Boolean


Private Sub chkIndex_Click()
'/* scheduler switch

    frmMain.chkIndex.Value = Me.chkIndex.Value
    If chkIndex.Value Then
        If chkIndex.Value Then
            Select Case True
            Case optIndex(0).Value
                Schedule_Add 1
            Case optIndex(1).Value
                Schedule_Add 2
            Case optIndex(2).Value
                Schedule_Add 3, cbDays.ListIndex
            End Select
        End If
    Else
        Schedule_Remove
    End If

End Sub

Private Sub chkStart_Click()

    frmMain.chkStart(2).Value = Me.chkStart.Value
    
End Sub

Private Sub cmdSet_Click(Index As Integer)

Dim sPath As String

    Select Case Index
    Case 0
    '/* start indexing
    sPath = left$(Drive1.Drive, 2) & Chr$(92)
    lblInfo(1).Visible = True
    lblInfo(1).Caption = "Status: Scanning drive " & sPath
    cmdSet(0).Enabled = False
    
    '/* store last indexed drive
    With New clsLightning
        .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "lstidx", Drive1.ListIndex
        .Write_String HKEY_CURRENT_USER, m_sRegPath, "drvpath", left$(Drive1.Drive, 2) & Chr$(92)
    End With
    
    '/* set save flag and start indexing
    With frmMain
        .Progress_Load
        .Index_Start sPath
    End With

    '/* unload
    Case 1
        cmdSet(0).Visible = True
        cmdSet(0).Enabled = True
        cmdSet(1).Visible = False
        frmMain.SetFocus
        Unload Me
    End Select
    
End Sub

Public Sub Drive_Reflect()

    m_bChange = True
    Me.Drive1.ListIndex = frmMain.Drive1.ListIndex
    Me.chkStart.Value = frmMain.chkStart(2).Value
    
End Sub

Private Sub Drive1_Change()

Dim sIndPath As String

On Error Resume Next

    '/* mirror drive changes
    If Not m_bChange Then
        frmMain.Drive_Reflect
    Else
        Get_Stats
        m_bChange = False
        Exit Sub
    End If

    sIndPath = App.Path & "\index" & CInt(Drive1.ListIndex) & ".dat"
    frmMain.txtPath.Text = left$(Drive1.Drive, 2) & Chr$(92)
    '/* loader switch
    If m_bChange Then Exit Sub

    '/* test for index, if exists, load
    '/* if not, invoke indexing dialog
    If File_Exists(sIndPath) Then
        cmdSet(0).Enabled = False
        lblInfo(1).Caption = "Index for Drive: " & left$(Drive1.Drive, 2) & Chr$(92) & " loaded."
        cmdSet(0).Visible = False
        cmdSet(1).Visible = True
    Else
        lblInfo(1).Caption = "Thers is no Index for the Drive: " & left$(Drive1.Drive, 2) & Chr$(92)
        Get_Stats
        cmdSet(0).Visible = True
        cmdSet(0).Enabled = True
        cmdSet(1).Visible = False
    End If

On Error GoTo 0

End Sub

Private Sub Get_Stats()

Dim lEstimate   As Long

    '/* get a rough scan time estimate
    '/* based on partition size, I will
    '/* add processor speed to equation lator
    m_lRefNum = Drive_Used(left$(Drive1.Drive, 2) & Chr$(92))
    lEstimate = m_lRefNum / 190
    lblStats(0).Caption = "Drive: " & left$(Drive1.Drive, 2) & Chr$(92)
    lblStats(1).Caption = "Used Space: " & m_lRefNum & " MegaBytes"
    lblStats(2).Caption = "Scan Time Est: " & lEstimate & " Seconds to Index"
    
End Sub

Private Sub Form_Load()

Dim i   As Integer

    lblInfo(1).Caption = "Please Select a Drive to Start.."
    m_lRefNum = Drive_Used(left$(Drive1.Drive, 2) & Chr$(92))
    
    Me.Drive1.ListIndex = frmMain.Drive1.ListIndex
    Me.chkIndex.Value = frmMain.chkIndex.Value
    Me.optIndex(0).Value = frmMain.optIndex(0).Value
    Me.optIndex(1).Value = frmMain.optIndex(1).Value
    Me.optIndex(2).Value = frmMain.optIndex(2).Value
    
    '/* sched days
    With cbDays
        .AddItem "Monday", 0
        .AddItem "Tuesday", 1
        .AddItem "Wednesday", 2
        .AddItem "Thursday", 3
        .AddItem "Friday", 4
        .AddItem "Saturday", 5
        .AddItem "Sunday", 6
        .ListIndex = 0
    End With
    
    Get_Stats

End Sub

Private Sub optIndex_Click(Index As Integer)

    If Not Forms = 1 Then
        frmMain.optIndex(Index).Value = True
    End If
    
End Sub
