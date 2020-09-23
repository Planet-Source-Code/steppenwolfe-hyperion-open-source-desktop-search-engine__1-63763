VERSION 5.00
Begin VB.Form frmMessage 
   Caption         =   "Alert"
   ClientHeight    =   2265
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4755
   Icon            =   "frmMessage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   2265
   ScaleWidth      =   4755
   StartUpPosition =   1  'CenterOwner
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   2265
      Left            =   0
      ScaleHeight     =   2235
      ScaleWidth      =   4725
      TabIndex        =   0
      Top             =   0
      Width           =   4755
      Begin VB.CommandButton cmdChoice 
         Caption         =   "Cancel"
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
         Index           =   2
         Left            =   3180
         TabIndex        =   5
         Top             =   1530
         Width           =   1185
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "No"
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
         Left            =   1740
         TabIndex        =   4
         Top             =   1530
         Width           =   1185
      End
      Begin VB.CommandButton cmdChoice 
         Caption         =   "Yes"
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
         Left            =   330
         TabIndex        =   3
         Top             =   1530
         Width           =   1185
      End
      Begin VB.Label lblHeader 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning!"
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
         Left            =   420
         TabIndex        =   2
         Top             =   180
         Width           =   3735
      End
      Begin VB.Label lblBody 
         BackStyle       =   0  'Transparent
         Caption         =   "Warning! "
         Height          =   915
         Left            =   420
         TabIndex        =   1
         Top             =   450
         Width           =   3825
      End
   End
End
Attribute VB_Name = "frmMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdChoice_Click(Index As Integer)

    Select Case Index
    Case 0
        m_iChoice = 1
    Case 1
        m_iChoice = 2
    Case 2
        m_iChoice = 3
    End Select
    Unload Me

End Sub

