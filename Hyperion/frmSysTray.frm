VERSION 5.00
Begin VB.Form frmSysTray 
   BorderStyle     =   0  'None
   ClientHeight    =   1575
   ClientLeft      =   6330
   ClientTop       =   5595
   ClientWidth     =   2055
   ControlBox      =   0   'False
   Icon            =   "frmSysTray.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1575
   ScaleWidth      =   2055
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   WindowState     =   1  'Minimized
   Begin VB.Menu mnuMain 
      Caption         =   ""
      Visible         =   0   'False
      Begin VB.Menu mnuMain1 
         Caption         =   "Open"
         Index           =   0
      End
      Begin VB.Menu mnuMain1 
         Caption         =   "About"
         Index           =   1
      End
      Begin VB.Menu mnuMain1 
         Caption         =   "-"
         Index           =   2
      End
      Begin VB.Menu mnuMain1 
         Caption         =   "Exit"
         Index           =   3
      End
   End
End
Attribute VB_Name = "frmSysTray"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'/* standard systray routines
Private Type NOTIFYICONDATA
    cbSize                            As Long
    hwnd                              As Long
    uId                               As Long
    uFlags                            As Long
    ucallbackMessage                  As Long
    hIcon                             As Long
    szTip                             As String * 64
End Type

Private Const NIM_ADD             As Long = &H0
Private Const NIM_DELETE          As Long = &H2
Private Const NIF_MESSAGE         As Long = &H1
Private Const NIF_ICON            As Long = &H2
Private Const NIF_TIP             As Long = &H4
Private Const WM_MOUSEMOVE        As Long = &H200
Private Const WM_LBUTTONDOWN      As Long = &H201
Private Const WM_LBUTTONUP        As Long = &H202
Private Const WM_LBUTTONDBLCLK    As Long = &H203
Private Const WM_RBUTTONDOWN      As Long = &H204
Private Const WM_RBUTTONUP        As Long = &H205
Private Const WM_RBUTTONDBLCLK    As Long = &H206

Private tpeNid                    As NOTIFYICONDATA

Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, _
                                                                                   pnid As NOTIFYICONDATA) As Boolean

Private Sub Form_Load()

    '/* show
    With Me
        .Show
        .Refresh
    End With

    '/* fill tray structure
    With tpeNid
        .cbSize = Len(tpeNid)
        .hwnd = Me.hwnd
        .uId = vbNull
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_MOUSEMOVE
        .hIcon = Me.Icon
        .szTip = " Right Click for Options" & vbNullChar
    End With

    Shell_NotifyIcon NIM_ADD, tpeNid

End Sub

Private Sub Form_MouseMove(Button As Integer, _
                           Shift As Integer, _
                           x As Single, _
                           y As Single)

Dim lReturn  As Long

    If Me.ScaleMode = vbPixels Then
        lReturn = x
    Else
        lReturn = x / Screen.TwipsPerPixelX
    End If

    '/* click responses
    Select Case lReturn
    Case WM_LBUTTONUP
        With Me
            .WindowState = vbNormal
            SetForegroundWindow .hwnd
            .Show
        End With
    Case WM_LBUTTONDBLCLK
        With Me
            .WindowState = vbNormal
            SetForegroundWindow .hwnd
            .Show
        End With
    Case WM_RBUTTONUP
        SetForegroundWindow Me.hwnd
        Me.PopupMenu Me.mnuMain
    End Select

End Sub

Private Sub Form_Resize()

    If Me.WindowState = vbMinimized Then
        Me.Hide
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Shell_NotifyIcon NIM_DELETE, tpeNid
    
    If frmMain.Visible = False Then
        With New clsLightning
            .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "loaded", 1
        End With
    End If
    
End Sub

Private Sub mnuMain1_Click(Index As Integer)
'/* menu options

    Select Case Index
    '/* open
    Case 0
        frmMain.Visible = True
    
    '/* about
    Case 1
        frmAbout.Visible = True
        
    '/* unload
    Case 3
       Unload_All
        
    End Select
    
End Sub
