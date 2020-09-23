VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About NSPowertools - Hyperion 1.1"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3270
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtAbout 
      Height          =   2415
      Left            =   150
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   150
      Width           =   4365
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
      Left            =   3330
      TabIndex        =   0
      Top             =   2700
      Width           =   1185
   End
   Begin VB.Label lblLink 
      AutoSize        =   -1  'True
      Caption         =   "Visit NSPowertools to learn more.."
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   150
      TabIndex        =   2
      Top             =   2880
      Width           =   2445
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*~ the subliminal thought of the day is  ~*

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub Form_Load()

    txtAbout.Text = "Hyperion, was the Sun Titan of Greek Legend, and was known as '..he who goes before the sun'. " & _
    "The software Hyperion also shines bright, as a tool of power and convenience, free, and easy to use, I have " & _
    "great expectations for this software. Hyperion is currently an Open Source project, which means I have decided " & _
    "to share both the finished product, and the original source code with the development world." & vbCrLf & "Hyperion uses our " & _
    "powerful Acchilles 1.4 Scan Engine, an extremely fast internal search engine, to power searches on your desktop at lightning speed." & _
     vbCrLf & _
    "Thanks for the puplished product go out to the legendary Steve McMahon for his excellent custom controls, and " & _
    "thanks also goes to Guenter Wirth for his wisdom and guidance during the creation of this software. Well, enjoy " & _
    "and visit us at our new location at www.nspowertools.com" & vbCrLf & vbCrLf & "John Underhill" & vbCrLf & _
    "Lead Developer - NSPowertools"
    
End Sub

Private Sub lblLink_Click()

    ShellExecute Me.hwnd, "open", "http://www.nspowertools.com/hyperion.htm", "", "", 1
    
End Sub
