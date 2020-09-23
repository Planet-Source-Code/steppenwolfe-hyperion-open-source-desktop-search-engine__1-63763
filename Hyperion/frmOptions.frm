VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmOptions 
   BackColor       =   &H00D8E9EC&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Options"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8925
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picPanel 
      Appearance      =   0  'Flat
      BackColor       =   &H00D8E9EC&
      ForeColor       =   &H80000008&
      Height          =   3315
      Index           =   0
      Left            =   210
      ScaleHeight     =   3285
      ScaleWidth      =   5475
      TabIndex        =   3
      Top             =   630
      Visible         =   0   'False
      Width           =   5505
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Write: wri"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   10
         Left            =   3330
         TabIndex        =   14
         Top             =   1650
         Width           =   1095
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Microsoft Works: wps"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   9
         Left            =   3330
         TabIndex        =   13
         Top             =   1335
         Width           =   1905
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Word for Mac: mcw"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   8
         Left            =   3330
         TabIndex        =   12
         Top             =   1020
         Width           =   1785
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Word: doc;dot"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   7
         Left            =   3330
         TabIndex        =   11
         Top             =   705
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Extended Text: ans"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   6
         Left            =   300
         TabIndex        =   10
         Top             =   2610
         Width           =   1845
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Rich Text: rtf"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   5
         Left            =   300
         TabIndex        =   9
         Top             =   2295
         Value           =   1  'Checked
         Width           =   1395
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Lotus 1-2-3: wk1;wk3;wk4"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   4
         Left            =   300
         TabIndex        =   8
         Top             =   1980
         Width           =   2265
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "HTML: htm;html;xhtm;xhtml;cfm"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   3
         Left            =   300
         TabIndex        =   7
         Top             =   1665
         Width           =   2655
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Excel: xls;xlw"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   6
         Top             =   1350
         Width           =   1335
      End
      Begin VB.CheckBox chkDFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "ASCII: txt"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1035
         Value           =   1  'Checked
         Width           =   1125
      End
      Begin VB.CheckBox chkDocFilter 
         BackColor       =   &H00D8E9EC&
         Caption         =   "Adobe Document: pdf"
         ForeColor       =   &H00404040&
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   4
         Top             =   720
         Value           =   1  'Checked
         Width           =   1935
      End
      Begin VB.Label lblHeader 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Document Search Filters"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   6
         Left            =   180
         TabIndex        =   15
         Top             =   210
         Width           =   2070
      End
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00D8E9EC&
      Caption         =   "Cancel"
      Height          =   345
      Index           =   1
      Left            =   4860
      TabIndex        =   2
      Top             =   4380
      Width           =   1035
   End
   Begin VB.CommandButton cmdOptions 
      BackColor       =   &H00D8E9EC&
      Caption         =   "OK"
      Height          =   345
      Index           =   0
      Left            =   3660
      TabIndex        =   1
      Top             =   4380
      Width           =   1035
   End
   Begin ComctlLib.TabStrip tbOptions 
      Height          =   4125
      Left            =   60
      TabIndex        =   0
      Top             =   60
      Width           =   5865
      _ExtentX        =   10345
      _ExtentY        =   7276
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   5
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Document Filters"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Image Filters"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Music Filters"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Exclusions Paths"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Advanced"
            Object.Tag             =   ""
            ImageVarType    =   2
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_Ival  As Integer

Private Sub Form_Load()

Dim x   As Integer

    For x = 1 To tbOptions.Index 'Loop through the tabs
        With picOptions(x)
            .BorderStyle = 0
            .Left = TabStrip1.ClientLeft
            .Top = TabStrip1.ClientTop
            .Width = TabStrip1.ClientWidth
            .Height = TabStrip1.ClientHeight
            .Visible = False
        End With
    Next x
    
    tbOptions.Tabs(1).Selected = True
    picOptions(TabStrip1.SelectedItem.Index).Visible = True 'Show first container
    
End Sub

Private Sub tbOptions_Click()

Dim x   As Integer

    m_Ival = Switch(m_Ival = 0, 1, m_Ival >= 1 And m_Ival <= Numtabs, m_Ival)
    picOptions(m_Ival).Visible = False
    picOptions(tbOptions.SelectedItem.Index).Visible = True
    picOptions(tbOptions.SelectedItem.Index).Refresh
    m_Ival = tbOptions.SelectedItem.Index
    
End Sub
