VERSION 5.00
Begin VB.Form frmSplash 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00FFFFFF&
   BorderStyle     =   0  'None
   ClientHeight    =   1905
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4155
   LinkTopic       =   "Form1"
   ScaleHeight     =   1905
   ScaleWidth      =   4155
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1905
      Left            =   0
      Picture         =   "frmSplash.frx":0000
      ScaleHeight     =   1960.217
      ScaleMode       =   0  'User
      ScaleWidth      =   4096.167
      TabIndex        =   0
      Top             =   0
      Width           =   4155
      Begin VB.Timer tmrLamer 
         Interval        =   5
         Left            =   240
         Top             =   480
      End
   End
End
Attribute VB_Name = "frmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//laser vars
Private m_iProg                  As Integer
Private m_iProg2                 As Integer
Private m_sLine2                 As String

'//laser const
Private Const HWND_TOPMOST     As Integer = -1
Private Const SWP_NOMOVE       As Long = &H2
Private Const SWP_NOSIZE       As Long = &H1
Private Const TOPMOST_FLAGS    As Double = SWP_NOMOVE Or SWP_NOSIZE

Private Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, _
                                                    ByVal hWndInsertAfter As Long, _
                                                    ByVal x As Long, _
                                                    ByVal y As Long, _
                                                    ByVal cX As Long, _
                                                    ByVal cY As Long, _
                                                    ByVal wFlags As Long) As Long

Private Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Private Sub Laser_Effect(picBox As PictureBox, _
                         ByVal StartX As Integer, _
                         ByVal StartY As Integer, _
                         ByVal iCount As Integer, _
                         ByVal iLine As Integer, _
                         ByVal sLine As String)

Dim lHeight As Long
Dim lWidth  As Long

    With Me
        .ScaleMode = vbPixels
        .Visible = True
        lHeight = .Height
        lWidth = .Width
        .Height = 0
        .Width = 0
    End With

    With picBox
        .ScaleMode = vbPixels
        .AutoRedraw = True
        .Visible = False
        
    End With

    For m_iProg2 = 0 To picBox.ScaleWidth
        For m_iProg = 0 To picBox.ScaleHeight
            m_sLine2 = picBox.Point(m_iProg2, m_iProg)
            Line (StartX, StartY)-(iCount + m_iProg2, iLine + m_iProg), m_sLine2
            DoEvents
        Next m_iProg
        Line (StartX, StartY)-(iCount + m_iProg2, iLine + picBox.ScaleHeight), sLine
        With Me
            If Not .Height > lHeight Then .Height = .Height + 20
            If Not .Width > lWidth Then .Width = .Width + 30
            'Sleep 1
        End With

        DoEvents
    Next m_iProg2

    For m_iProg2 = 0 To picBox.ScaleHeight
        Line (StartX, StartY)-(iCount + picBox.ScaleWidth, iLine + m_iProg2), sLine
    Next m_iProg2

End Sub

Private Sub tmrLamer_Timer()

    Laser_Effect Picture1, 400, 350, 0, 0, Me.BackColor
    If Forms > 1 Then
    tmrLamer.Enabled = False
        Unload Me
    End If

End Sub

Private Sub Form_Load()

    On Error Resume Next

    tmrLamer.Enabled = False
    Me.Move (Screen.Width - Width) \ 2, ((Screen.Height - Height) \ 2) - 500
    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS

    On Error GoTo 0

End Sub
