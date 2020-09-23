VERSION 5.00
Begin VB.Form frmAlert 
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1590
   ClientLeft      =   5745
   ClientTop       =   3870
   ClientWidth     =   2280
   ControlBox      =   0   'False
   Icon            =   "frmAlert.frx":0000
   LinkTopic       =   "frmAlert"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   106
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   152
   ShowInTaskbar   =   0   'False
   Visible         =   0   'False
   Begin VB.Timer Timer1 
      Interval        =   1
      Left            =   2925
      Top             =   2250
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      ForeColor       =   &H00404040&
      Height          =   195
      Left            =   -15
      TabIndex        =   0
      Top             =   0
      UseMnemonic     =   0   'False
      Width           =   2205
      WordWrap        =   -1  'True
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'*~ Morgan Haueisen's awesome alert window.. ~*
'*~ http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=52793&lngWId=1 ~*
'*~ If you write user controls, this is a good example of proffesional programming.. ~*

'/* A Project Global variable is required for displaying more then 1 at a time
'/* See Sub GetDisplayPosition for more information

'/* Available screen size (without task bar)

Private Const SPI_GETWORKAREA                As Long = 48&
Private Type Rect
    left                                     As Long
    top                                      As Long
    right                                    As Long
    bottom                                   As Long
End Type

'/* GradientFill API - Requires Windows 2000 or later; Requires Windows 98 or later
Private Type GRADIENT_TRIANGLE
    Vertex1                                  As Long
    Vertex2                                  As Long
    Vertex3                                  As Long
End Type

Private Type GRADIENT_RECT
    UpperLeft                                As Long    '/* UNSIGNED Long
    LowerRight                               As Long    '/* UNSIGNED Long
End Type

Private Type TRIVERTEX
    x                                        As Long
    y                                        As Long
    Red                                      As Integer '/* Ushort value
    Green                                    As Integer '/* Ushort value
    Blue                                     As Integer '/* Ushort value
    Alpha                                    As Integer '/* Ushort value
End Type

Private Const GRADIENT_FILL_TRIANGLE        As Long = &H2&
Private Const GRADIENT_FILL_RECT_V          As Long = &H1&
Private Const LWA_ALPHA                     As Long = &H2
Private Const GWL_EXSTYLE                   As Long = (-20)
Private Const WS_EX_LAYERED                 As Long = &H80000

'/* Operating system version information
Private Type OSVersionInfo
    OSVSize                                  As Long
    dwVerMajor                               As Long
    dwVerMinor                               As Long
    dwBuildNumber                            As Long
    PlatformID                               As Long
    szCSDVersion                             As String * 128
End Type

'/* Used to draw the form's rounded border
'/* Used to make the rounded corners of the form transparent
'/* Form variables
Private m_lChangeSpeed                      As Long    '/* The window's display speed
Private m_lCounter                          As Long    '/* Display time in milliseconds
Private m_lScrnBottom                       As Long
    '/* Height of the screen - taskbar (if it is on the bottom)
Private m_bOnTop                            As Boolean '/* Form Z-Order Flag
Private m_lWindowCount                      As Long
    '/* Screen stop position multiplier (displaying more then 1 at a time)
Private m_bManualClose                      As Boolean '/* Manual close Flag
Private m_bCodeClose                        As Boolean '/* Prevent user close option
Private m_bFade                             As Boolean '/* Fade or move Flag
Private m_iOSver                            As Byte    '/* OS 1=Win98/ME; 2=Win2000/XP
Private m_lActivate                         As Long

Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, _
                                                                                          ByVal uParam As Long, _
                                                                                          ByRef lpvParam As Any, _
                                                                                          ByVal fuWinIni As Long) As Long

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cX As Long, _
                                                    ByVal cY As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Function GradientFillTriangle Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, _
                                                                                  pVertex As TRIVERTEX, _
                                                                                  ByVal dwNumVertex As Long, _
                                                                                  pMesh As GRADIENT_TRIANGLE, _
                                                                                  ByVal dwNumMesh As Long, _
                                                                                  ByVal dwMode As Long) As Long

Private Declare Function GradientFillRect Lib "msimg32" Alias "GradientFill" (ByVal hdc As Long, _
                                                                              pVertex As TRIVERTEX, _
                                                                              ByVal dwNumVertex As Long, _
                                                                              pMesh As GRADIENT_RECT, _
                                                                              ByVal dwNumMesh As Long, _
                                                                              ByVal dwMode As Long) As Long

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long) As Long

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, _
                                                                            ByVal nIndex As Long, _
                                                                            ByVal dwNewLong As Long) As Long

Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, _
                                                                  ByVal crKey As Long, _
                                                                  ByVal bAlpha As Byte, _
                                                                  ByVal dwFlags As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVersionInfo) As Long

Private Declare Function RoundRect Lib "gdi32" (ByVal hdc As Long, _
                                                ByVal left As Long, _
                                                ByVal top As Long, _
                                                ByVal right As Long, _
                                                ByVal bottom As Long, _
                                                ByVal EllipseWidth As Long, _
                                                ByVal EllipseHeight As Long) As Long

Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hRgn As Long, _
                                                    ByVal bRedraw As Long) As Long

Private Declare Function CreateRoundRectRgn Lib "gdi32" (ByVal RectX1 As Long, _
                                                         ByVal RectY1 As Long, _
                                                         ByVal RectX2 As Long, _
                                                         ByVal RectY2 As Long, _
                                                         ByVal EllipseWidth As Long, _
                                                         ByVal EllipseHeight As Long) As Long

Private Sub Form_Load()

Dim Rc        As Rect
Dim scrnRight As Long
Dim OSV       As OSVersionInfo

    OSV.OSVSize = Len(OSV)
    If GetVersionEx(OSV) = 1 Then
        If OSV.PlatformID = 1 Then
            If OSV.dwVerMinor >= 10 Then
                m_iOSver = 1
            End If
        End If
        If OSV.PlatformID = 2 Then
            If OSV.dwVerMajor >= 5 Then
                m_iOSver = 2
            End If
        End If
    End If

    '/* Get Screen and TaskBar size
    SystemParametersInfo SPI_GETWORKAREA, 0&, Rc, 0&
    '/* Screen Height - Taskbar Height (if is is located at the bottom of the screen)
    m_lScrnBottom = Rc.bottom * Screen.TwipsPerPixelY
    '/* Is the taskbar is located on the right side of the screen? (scrnRight < Screen.width)
    scrnRight = (Rc.right * Screen.TwipsPerPixelX)
    '/* Locate Form to bottom right and set default size
    Me.Move scrnRight - 2400, m_lScrnBottom, 2400, 2000
    GetDisplayPosition True

End Sub

Private Sub Form_Click()

'/* Close by user
    If Not m_bCodeClose Then
        m_lCounter = 0
    End If

End Sub

Private Sub lblMessage_Click()

'/* Close by user
    If Not m_bCodeClose Then
        m_lCounter = 0
    End If

End Sub

Public Sub DisplayMessage(ByVal sMessage As String, _
                          Optional ByVal DisplaySeconds As Long = 4, _
                          Optional ByVal bFade As Boolean = False, _
                          Optional ByVal bAutoFit As Boolean = True, _
                          Optional ByVal bSquare As Boolean = True, _
                          Optional ByVal iBackColor As Long = &HC0FFFF, _
                          Optional ByVal bTubeFill As Boolean = False)

    If DisplaySeconds < 1& Then '/* Manual Close
        m_bManualClose = True
        m_lCounter = 1&
        If DisplaySeconds = 0& Then
            sMessage = "(click here to close)" & vbNewLine & sMessage
        Else '/* DisplaySeconds < 0
            '/* Close by code only
            m_bCodeClose = True
        End If
    Else '/* Auto Close
        '/* Convert seconds to milliseconds
        m_lCounter = DisplaySeconds * 100&
    End If

    Me.ScaleMode = vbPixels

    '/* Resize the Form's height based on the amount of text to display
    '/* If more then one alert is showing then fix the height to standard to insure no overlap
    If m_lCounter = 1 Then
        lblMessage.Move 5, 5, Me.ScaleWidth - 10
    Else
        lblMessage.Move 5, 10, Me.ScaleWidth - 10
    End If
    lblMessage.Caption = sMessage
    If m_lWindowCount = 1 And bAutoFit Then
        Me.Height = (lblMessage.top + lblMessage.Height + 15) * Screen.TwipsPerPixelY
    End If

    '/* Move or Fade?
    m_bFade = bFade
    If m_bFade Then
        '/* Start with 100% transparent
        m_lChangeSpeed = 100&
        Me.top = m_lScrnBottom - (Me.Height * m_lWindowCount)
        '/* prevent it from going over the top of the screen
        If Me.top < 0 Then
            Me.top = 0
        End If
        MakeTransparent m_lChangeSpeed
        SetOnTop True
    Else
        '/* Move distance per millisecond
        If m_lWindowCount > 1 Then
            m_lChangeSpeed = 100&
        Else
            m_lChangeSpeed = 50&
        End If
    End If

    '/* Add colored background
    If bTubeFill Then
        GradientFillTube iBackColor
    Else
        GradientFill iBackColor
    End If

    If bSquare Then
        '/* Draw Square borders around the Form
        Me.Line (Me.ScaleWidth - 1, Me.ScaleHeight - 1)-(Me.ScaleWidth - 1, 0), vbButtonFace
        Me.Line (Me.ScaleWidth - 1, Me.ScaleHeight - 1)-(0, Me.ScaleHeight - 1), vbButtonFace
        Me.Line (Me.ScaleWidth - 1, 0)-(0, 0), vbButtonFace
        Me.Line (0, Me.ScaleHeight - 1)-(0, 0), vbButtonFace
        Me.Line (Me.ScaleWidth - 2, Me.ScaleHeight - 2)-(Me.ScaleWidth - 2, 1), iBackColor
        Me.Line (Me.ScaleWidth - 2, Me.ScaleHeight - 2)-(1, Me.ScaleHeight - 2), iBackColor
        Me.Line (Me.ScaleWidth - 2, 1)-(1, 1), iBackColor
        Me.Line (1, Me.ScaleHeight - 2)-(1, 1), iBackColor
    Else
        '/* Draw rounded borders around the Form
        With Me
            .ForeColor = vbButtonFace
            RoundRect .hdc, 0&, 0&, .ScaleWidth - 1, .ScaleHeight - 1, 20&, 20&
            .ForeColor = iBackColor
            RoundRect .hdc, 1, 1, .ScaleWidth - 2, .ScaleHeight - 2, 18&, 18&
            '/* Make corners transparent
            SetWindowRgn .hwnd, CreateRoundRectRgn(0&, 0&, .ScaleWidth, .ScaleHeight, 19&, 19&), True
        End With
    End If

    '/* Make sure the form is visible
    Me.Show
    '/* Begin - this could be done without a timer control (which is interrupt driven)
    '/* but it would be very demanding on CPU process time.
    Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()

    If m_bFade Then '/* Fade

        If m_lCounter > 0 Then '/* Fade In then Wait
            If m_lChangeSpeed > 0 Then
                m_lChangeSpeed = m_lChangeSpeed - 2
                MakeTransparent m_lChangeSpeed
            Else
                '/* Wait
                If Not m_bManualClose Then
                    m_lCounter = m_lCounter - 1
                End If
            End If

        Else '/* Fade Out then Close
            If m_lChangeSpeed <= 100 Then
                '/* Fade out
                m_lChangeSpeed = m_lChangeSpeed + 2
                MakeTransparent m_lChangeSpeed
            Else
                '/* Close
                Unload Me
            End If
        End If
    Else '/* Move

        If m_lCounter > 0 Then '/* Move Up then Wait
            If Me.top > 0 And Me.top > m_lScrnBottom - (Me.Height * m_lWindowCount) Then
                '/* Move Up
                Me.top = Me.top - m_lChangeSpeed
            Else
                '/* Wait
                If Not m_bOnTop Then
                    Me.top = m_lScrnBottom - (Me.Height * m_lWindowCount)
                    '/* prevent it from going over the top of the screen
                    If Me.top < 0 Then
                        Me.top = 0
                    End If
                    m_bOnTop = True
                    SetOnTop m_bOnTop
                End If
                If Not m_bManualClose Then
                    m_lCounter = m_lCounter - 1
                End If
            End If

        Else '/* Move Down then Close
            If Me.top <= Screen.Height Then
                '/* Move Down
                If m_bOnTop Then
                    m_bOnTop = False
                    SetOnTop m_bOnTop
                End If
                Me.top = Me.top + m_lChangeSpeed
            Else
                '/* Close
                Unload Me
            End If
        End If

    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    GetDisplayPosition False
    '/* Clean-up and clear memory
    Set frmAlert = Nothing

End Sub

Private Sub SetOnTop(Optional ByVal bSetOnTop As Boolean = True)

'/* The SetWindowPos function changes the size, position, and Z order of a child,
'/* pop-up, or top-level window. Child, pop-up, and top-level windows are ordered
'/* according to their appearance on the screen. The topmost window receives the
'/* highest rank and is the first window in the Z order.

Const Flags As Long = &H273

    '/* SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE or SWP_NOOWNERZORDER Or SWP_FRAMECHANGED

    If bSetOnTop Then
        SetWindowPos Me.hwnd, -1, 0, 0, 0, 0, Flags
    Else
        SetWindowPos Me.hwnd, -2, 0, 0, 0, 0, Flags
    End If

End Sub

Private Sub GradientFill(ByVal iBackColor As Long)

Dim TriVert(3) As TRIVERTEX
Dim gTRi(1)    As GRADIENT_TRIANGLE

    '/* Requires Windows 2000 or later; Requires Windows 98/ME
    If m_iOSver = 0 Then
        Me.BackColor = iBackColor
    Else
        Me.AutoRedraw = True
        'Me.ScaleMode = vbPixels '/* Required but done in Sub DisplayMessage

        '/* Top Left Trangle
        TriVert(0).x = 0&
        TriVert(0).y = 0&
        GradientFillColor TriVert(0), &HCCCCCC

        '/* Top Right Trangle
        TriVert(1).x = Me.ScaleWidth
        TriVert(1).y = 0&
        GradientFillColor TriVert(1), &HCCCCCC

        '/* Bottom Right Trangle
        TriVert(2).x = Me.ScaleWidth
        TriVert(2).y = Me.ScaleHeight
        GradientFillColor TriVert(2), iBackColor

        '/* Bottom Left Trangle
        TriVert(3).x = 0&
        TriVert(3).y = Me.ScaleHeight
        GradientFillColor TriVert(3), iBackColor

        With gTRi(0)
            .Vertex1 = 0&
            .Vertex2 = 1&
            .Vertex3 = 2&
        End With 'gTRi(0)

        With gTRi(1)
            .Vertex1 = 0&
            .Vertex2 = 2&
            .Vertex3 = 3&

        End With 'gTRi(1)
        GradientFillTriangle Me.hdc, TriVert(0), 4&, gTRi(0), 2&, GRADIENT_FILL_TRIANGLE

    End If

End Sub

Private Sub GradientFillColor(ByRef tTV As TRIVERTEX, _
                              ByVal iColor As Long)

Dim iRed   As Long
Dim iGreen As Long
Dim iBlue  As Long


    '/* Separate color into RGB
    iRed = (iColor And &HFF&) * &H100&
    iGreen = (iColor And &HFF00&)
    iBlue = (iColor And &HFF0000) \ &H100&

    '/* Make Red color a UShort
    If (iRed And &H8000&) = &H8000& Then
        tTV.Red = (iRed And &H7F00&)
        tTV.Red = tTV.Red Or &H8000
    Else
        tTV.Red = iRed
    End If
    '/* Make Green color a UShort
    If (iGreen And &H8000&) = &H8000& Then
        tTV.Green = (iGreen And &H7F00&)
        tTV.Green = tTV.Green Or &H8000
    Else
        tTV.Green = iGreen
    End If
    '/* Make Blue color a UShort
    If (iBlue And &H8000&) = &H8000& Then
        tTV.Blue = (iBlue And &H7F00&)
        tTV.Blue = tTV.Blue Or &H8000
    Else
        tTV.Blue = iBlue
    End If

End Sub

Private Sub MakeTransparent(ByVal PercentTransparent As Long)

Dim Ret As Long

    '/* Requires Windows 2000 or later; Win9x/ME is Not supported
    If m_iOSver = 2 Then
On Error Resume Next
        '/* Convert 0-100 to 255-0
        PercentTransparent = ((100& - PercentTransparent) / 100&) * 255&
        If PercentTransparent >= 0& Then
            If PercentTransparent <= 255& Then
                Ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
                Ret = Ret Or WS_EX_LAYERED
                SetWindowLong Me.hwnd, GWL_EXSTYLE, Ret
                SetLayeredWindowAttributes Me.hwnd, 0&, PercentTransparent, LWA_ALPHA
            End If
        End If

    End If

On Error GoTo 0

End Sub

Private Sub GetDisplayPosition(ByVal SetPosition As Boolean)
    
    '/* Project Global variable m_lActivate) is required for displaying more then 1 at a time
    '/* This option can be removed by always setting m_lWindowCount = 1 and commenting out the rest of this sub
    
    '/* This binary addition is required to insure that a newly created window does not cover
    '/* a window that is already showing. I stopped at 8 but you could go to 15 if m_lActivate
    '/* is defined as an Integer or 31 if it is defined as a Long
    
    If SetPosition Then
        '/* Reserve a window position
        If (m_lActivate And 1) = 0 Then
            m_lWindowCount = 1
            m_lActivate = m_lActivate Or 1
        ElseIf (m_lActivate And 2) = 0 Then
            m_lWindowCount = 2
            m_lActivate = m_lActivate Or 2
        ElseIf (m_lActivate And 4) = 0 Then
            m_lWindowCount = 3
            m_lActivate = m_lActivate Or 4
        ElseIf (m_lActivate And 8) = 0 Then
            m_lWindowCount = 4
            m_lActivate = m_lActivate Or 8
        ElseIf (m_lActivate And 16) = 0 Then
            m_lWindowCount = 5
            m_lActivate = m_lActivate Or 16
        ElseIf (m_lActivate And 32) = 0 Then
            m_lWindowCount = 6
            m_lActivate = m_lActivate Or 32
        ElseIf (m_lActivate And 64) = 0 Then
            m_lWindowCount = 7
            m_lActivate = m_lActivate Or 64
        Else
            m_lWindowCount = 8
            m_lActivate = m_lActivate Or 128
        End If
    
    Else
        '/* Free up window position for use
        Select Case m_lWindowCount
        Case 1
            m_lActivate = m_lActivate Xor 1
        Case 2
            m_lActivate = m_lActivate Xor 2
        Case 3
            m_lActivate = m_lActivate Xor 4
        Case 4
            m_lActivate = m_lActivate Xor 8
        Case 5
            m_lActivate = m_lActivate Xor 16
        Case 6
            m_lActivate = m_lActivate Xor 32
        Case 7
            m_lActivate = m_lActivate Xor 64
        Case Else
            m_lActivate = m_lActivate Xor 128
        End Select
    End If
  
End Sub

Private Sub GradientFillTube(ByVal iBackColor As Long)

    Dim TriVert(1) As TRIVERTEX
    Dim gRect      As GRADIENT_RECT

    '/* Requires Windows 2000 or later; Requires Windows 98/ME
    If m_iOSver = 0 Then
        Me.BackColor = iBackColor
        Exit Sub
    End If
    
    On Error Resume Next
    Me.AutoRedraw = True
    'Me.ScaleMode = vbPixels '/* Required but done in Sub DisplayMessage

    gRect.UpperLeft = 1
    gRect.LowerRight = 0
    
    '/* Top to Bottom
    '/* Draw top half
    With TriVert(0)
        .x = 0
        .y = 0
    End With
    Call GradientFillColor(TriVert(0), iBackColor)

    With TriVert(1)
        .x = Me.ScaleWidth
        .y = Me.ScaleHeight \ 2
    End With
    Call GradientFillColor(TriVert(1), vbWhite)
    Call GradientFillRect(Me.hdc, TriVert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)
    
    '/* Draw bottom half
    With TriVert(0)
        .x = 0
        .y = Me.ScaleHeight \ 2
    End With
    Call GradientFillColor(TriVert(0), vbWhite)

    With TriVert(1)
        .x = Me.ScaleWidth
        .y = Me.ScaleHeight
    End With
    Call GradientFillColor(TriVert(1), iBackColor)
    
    Call GradientFillRect(Me.hdc, TriVert(0), 2, gRect, 1, GRADIENT_FILL_RECT_V)

End Sub

Public Property Get CloseActivate() As Boolean
    '/* Not Used
End Property

Public Property Let CloseActivate(ByVal vNewValue As Boolean)
    '/* Close Form from code
    If vNewValue Then m_lCounter = 0
End Property

