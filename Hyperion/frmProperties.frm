VERSION 5.00
Begin VB.Form frmProperties 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "File Properties"
   ClientHeight    =   2895
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5445
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmProperties.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2895
   ScaleWidth      =   5445
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   0
      ScaleHeight     =   2865
      ScaleWidth      =   5415
      TabIndex        =   0
      Top             =   0
      Width           =   5445
      Begin VB.PictureBox Pic32 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   480
         Left            =   4740
         ScaleHeight     =   480
         ScaleWidth      =   480
         TabIndex        =   6
         Top             =   210
         Width           =   480
      End
      Begin VB.CommandButton cmdExit 
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
         Left            =   4020
         TabIndex        =   2
         Top             =   2280
         Width           =   1185
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Detail:"
         Height          =   195
         Index           =   5
         Left            =   120
         TabIndex        =   8
         Top             =   2010
         Width           =   465
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Ver:"
         Height          =   195
         Index           =   4
         Left            =   120
         TabIndex        =   7
         Top             =   1710
         Width           =   300
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Desc:"
         Height          =   195
         Index           =   3
         Left            =   120
         TabIndex        =   5
         Top             =   1410
         Width           =   405
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Size:"
         Height          =   195
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1140
         Width           =   345
      End
      Begin VB.Label lblProperties 
         Caption         =   "Path:"
         Height          =   195
         Index           =   1
         Left            =   120
         TabIndex        =   3
         Top             =   840
         Width           =   5010
      End
      Begin VB.Label lblProperties 
         AutoSize        =   -1  'True
         Caption         =   "Name:"
         Height          =   195
         Index           =   0
         Left            =   120
         TabIndex        =   1
         Top             =   570
         Width           =   465
      End
   End
End
Attribute VB_Name = "frmProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/* icon constants
Private Const MAX_PATH                As Integer = 260
Private Const ILD_TRANSPARENT               As Long = &H1
Private Const SHGFI_DISPLAYNAME             As Long = &H200
Private Const SHGFI_EXETYPE                 As Long = &H2000
Private Const SHGFI_SYSICONINDEX            As Long = &H4000
Private Const SHGFI_LARGEICON               As Long = &H0
Private Const SHGFI_SHELLICONSIZE           As Long = &H4
Private Const SHGFI_TYPENAME                As Long = &H400
Private Const BASIC_SHGFI_FLAGS             As Double = SHGFI_TYPENAME Or _
SHGFI_SHELLICONSIZE Or SHGFI_SYSICONINDEX Or SHGFI_DISPLAYNAME Or SHGFI_EXETYPE

Private Type SHFILEINFO
    hIcon                                As Long
    iIcon                                As Long
    dwAttributes                         As Long
    szDisplayName                        As String * MAX_PATH
    szTypeName                           As String * 80
End Type

Private ShInfo                                  As SHFILEINFO

Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, _
                                                                                 ByVal dwFileAttributes As Long, _
                                                                                 psfi As SHFILEINFO, _
                                                                                 ByVal cbSizeFileInfo As Long, _
                                                                                 ByVal uFlags As Long) As Long

Private Declare Function ImageList_Draw Lib "comctl32.dll" (ByVal himl As Long, _
                                                            ByVal i As Long, _
                                                            ByVal hDCDest As Long, _
                                                            ByVal x As Long, _
                                                            ByVal y As Long, _
                                                            ByVal Flags As Long) As Long

Private Sub cmdExit_Click()
'/* exit

    Unload Me
    frmMain.SetFocus
    
End Sub

Public Sub File_Data(ByVal sPath As String, _
                     ByVal sSize As String)
'/* fetch file data

Dim lIcon   As Long

    '/* initiate class and interrogate file
    With New clsProperties
        .FileName = sPath
        lblProperties(0).Caption = "Name: " & Mid$(sPath, InStrRev(sPath, Chr$(92)) + 1)
        lblProperties(1).Caption = "Path: " & sPath
        lblProperties(2).Caption = "Size: " & sSize
        If Len(.Comments) = 0 Then
            lblProperties(3).Caption = "Desc: Unknown"
        Else
            lblProperties(3).Caption = "Desc: " & .Comments
        End If
        
        If Len(.ProductVersion) = 0 Then
            lblProperties(4).Caption = "Ver: Unknown"
        Else
            lblProperties(4).Caption = "Ver: " & .ProductVersion
        End If
        
        If Len(.ProductName) = 0 Then
            lblProperties(5).Caption = "Detail: Unknown"
        Else
            lblProperties(5).Caption = "Detail: " & .ProductName
        End If
    End With
    
    '/* get the icon associated with the file
    lIcon = SHGetFileInfo(sPath, 0&, ShInfo, Len(ShInfo), BASIC_SHGFI_FLAGS Or SHGFI_LARGEICON)
    If lIcon = 0 Then Exit Sub
    With Pic32
        Set .Picture = LoadPicture("")
        .AutoRedraw = True
        ImageList_Draw lIcon, ShInfo.iIcon, .hdc, 0, 0, ILD_TRANSPARENT
        .Refresh
    End With
    
End Sub
