Attribute VB_Name = "mMain"
Option Explicit

Public m_sRegPath       As String
Public m_sAppPath       As String
Public m_bWin32         As Boolean
Public m_iChoice        As Integer
Public cCounter         As New Collection

Private Const VER_PLATFORM_WIN32s             As Integer = 0
Private Const VER_PLATFORM_WIN32_WINDOWS      As Integer = 1
Private Const VER_PLATFORM_WIN32_NT           As Integer = 2

'/* version structure
Private Type OSVersion
    dwOSVersionInfoSize                           As Long
    dwMajorVersion                                As Long
    dwMinorVersion                                As Long
    dwBuildNumber                                 As Long
    dwPlatformId                                  As Long
    szCSDVersion                                  As String * 128
End Type

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersion As OSVersion) As Boolean


Public Sub Main()
'/* startup options

'/* reg flag: loaded
'/* normal (on start)
'/* silent (index update)
'/* loaded switch: 1=unloaded/2=running/3=minimized

Dim sPath       As String
Dim lLoaded     As Long
Dim lIndex      As Long
Dim bIndexState As Boolean

On Error Resume Next

    '/* set col and os id
    InitCommonControls
    Identify_OS
    Load frmMain
    
    '/* reg/app path globals
    m_sRegPath = "Software\" & App.ProductName
    m_sAppPath = App.Path & Chr$(92) & App.EXEName & ".exe"
    
    '/* get running state
    With New clsLightning
        lLoaded = .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "loaded")
    End With
    
    Select Case LCase$(Command)
    
    '/* silent update
    Case "/s", "-s"
        '/* if app is live, abort and set todo flag
        If lLoaded = 2 Then
            With New clsLightning
                .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "schreset", 1
            End With
            End
        Else
            '/* get the drive index path
            With New clsLightning
                sPath = .Read_String(HKEY_CURRENT_USER, m_sRegPath, "drvpath")
                .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "schreset", 0
            End With
            
            '/* test value
            If Len(sPath) = 0 Then End
            
            '/* enumerate and update indexes
            frmMain.Index_Control 3
        End If
        
    '/* run on startup
    Case "/r", "-r"
        '/* test index status and load accordingly
        frmSplash.Show
        DoEvents
        If frmMain.Index_Control(1) Then
            Unload frmSplash
            frmMain.Show
        Else
            frmMain.Show
            frmMain.Drive_Reflect
            frmQuery.Show vbModeless, frmMain
            Unload frmSplash
        End If
        
    '/* user launched
    Case Else
        '/* test index status and load accordingly
        frmSplash.Show
        DoEvents
        If frmMain.Index_Control(1) Then
            Unload frmSplash
            frmMain.Show
        Else
            frmMain.Show
            frmMain.Drive_Reflect
            frmQuery.Show vbModeless, frmMain
            Unload frmSplash
        End If
        
    End Select

On Error GoTo 0

End Sub

Public Sub Identify_OS()
'/* set os version flag

Dim rOsVersion As OSVersion

    rOsVersion.dwOSVersionInfoSize = Len(rOsVersion)
    If GetVersionEx(rOsVersion) Then
        If Not rOsVersion.dwPlatformId = VER_PLATFORM_WIN32_NT Then
            m_bWin32 = True
        End If
    End If

End Sub

