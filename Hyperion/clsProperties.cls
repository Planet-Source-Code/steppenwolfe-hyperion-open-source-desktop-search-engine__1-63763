VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsProperties"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'/* as taken (and marginally rewritten) from a great example on the scarms.com
'/* sorry, took all the comments out a long time ago.., check out the original
'/* if you want to know how it works..

Private Const adhcClassName                 As String = "VersionInfo"
Private Const adhcErrBase                   As Double = vbObjectError + 19560516
Private Const adhcErrNoModuleHandle         As Double = adhcErrBase + 1
Private Const adhcErrNoModule               As Double = adhcErrBase + 2
Private Const adhcErrNoFileName             As Double = adhcErrBase + 3
Private Const adhcErrNoVersionInfoSize      As Double = adhcErrBase + 4
Private Const adhcErrNoVersionInfo          As Double = adhcErrBase + 5
Private Const adhcErrNoLangCPInfo           As Double = adhcErrBase + 6
Private Const adhcNoFixedFileInfo           As Double = adhcErrBase + 7
Private Const adhcErrUnknown                As Double = adhcErrBase + 8

'/* language enum
Public Enum adhLanguageValue
    lArabic = 1025
    lBulgarian = 1026
    lCatalan = 1027
    lTraditionalChinese = 1028
    lCzech = 1029
    lDanish = 1030
    lGerman = 1031
    lGreek = 1032
    lUSEnglish = 1033
    lCastilianSpanish = 1034
    lFinnish = 1035
    lFrench = 1036
    lHebrew = 1037
    lHungarian = 1038
    lIcelandic = 1039
    lItalian = 1040
    lJapanese = 1041
    lKorean = 1042
    lDutch = 1043
    lNorwegianBokmål = 1044
    lPolish = 1045
    lBrazilianPortuguese = 1046
    lRhaetoRomanic = 1047
    lRomanian = 1048
    lRussian = 1049
    lCroatoSerbian = 1050
    lSlovak = 1051
    lAlbanian = 1052
    lSwedish = 1053
    lThai = 1054
    lTurkish = 1055
    lUrdu = 1056
    lBahasa = 1057
    lSimplifiedChinese = 2052
    lSwissGerman = 2055
    lUKEnglish = 2057
    lMexicanSpanish = 2058
    lBelgianFrench = 2060
    lSwissItalian = 2064
    lBelgianDutch = 2067
    lNorwegianNynorsk = 2068
    lPortuguese = 2070
    lSerboCroatian = 2074
    lCanadianFrench = 3084
    lSwissFrench = 4108
End Enum

'/* file info structure
Private Type VS_FIXEDFILEINFO
    dwSignature                                As Long
    dwStrucVersionl                            As Integer
    dwStrucVersionh                            As Integer
    dwFileVersionMSl                           As Integer
    dwFileVersionMSh                           As Integer
    dwFileVersionLSl                           As Integer
    dwFileVersionLSh                           As Integer
    dwProductVersionMSl                        As Integer
    dwProductVersionMSh                        As Integer
    dwProductVersionLSl                        As Integer
    dwProductVersionLSh                        As Integer
    dwFileFlagsMask                            As Long
    dwFileFlags                                As Long
    dwFileOS                                   As Long
    dwFileType                                 As Long
    dwFileSubtype                              As Long
    dwFileDateMS                               As Long
    dwFileDateLS                               As Long
End Type

'/* file flags
Public Enum adhFileFlags
    VS_FF_DEBUG = &H1&
    VS_FF_PRERELEASE = &H2&
    VS_FF_PATCHED = &H4&
    VS_FF_PRIVATEBUILD = &H8&
    VS_FF_SPECIALBUILD = &H20&
End Enum

'/* os type
Public Enum adhTargetOS
    VOS_UNKNOWN = &H0&
    VOS__BASE = &H0&
    VOS__WINDOWS16 = &H1&
    VOS__PM16 = &H2&
    VOS__PM32 = &H3&
    VOS__WINDOWS32 = &H4&
    VOS_DOS = &H10000
    VOS_OS216 = &H20000
    VOS_OS232 = &H30000
    VOS_NT = &H40000
    VOS_DOS_WINDOWS16 = &H10001
    VOS_DOS_WINDOWS32 = &H10004
    VOS_NT_WINDOWS32 = &H40004
    VOS_OS216_PM16 = &H20002
    VOS_OS232_PM32 = &H30003
End Enum

'/* file type
Public Enum adhFileType
    VFT_UNKNOWN = &H0&
    VFT_APP = &H1&
    VFT_DLL = &H2&
    VFT_DRV = &H3&
    VFT_FONT = &H4&
    VFT_VXD = &H5&
    VFT_STATIC_LIB = &H7&
End Enum

Public Enum adhFileSubType
    VFT2_UNKNOWN = &H0&
    VFT2_DRV_COMM = &HA&
    VFT2_DRV_PRINTER = &H1&
    VFT2_DRV_KEYBOARD = &H2&
    VFT2_DRV_LANGUAGE = &H3&
    VFT2_DRV_DISPLAY = &H4&
    VFT2_DRV_MOUSE = &H5&
    VFT2_DRV_NETWORK = &H6&
    VFT2_DRV_SYSTEM = &H7&
    VFT2_DRV_INSTALLABLE = &H8&
    VFT2_DRV_SOUND = &H9&
    VFT2_DRV_INPUTMETHOD = &HB&

    VFT2_FONT_RASTER = &H1&
    VFT2_FONT_TRUETYPE = &H3&
    VFT2_FONT_VECTOR = &H2&
End Enum

Private mstrModuleName                          As String
Private mstrFileName                            As String
Private mabytBuffer()                           As Byte
Private mstrLangCP                              As String
Private mffi                                    As VS_FIXEDFILEINFO

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, _
                                                                                      ByVal lpFileName As String, _
                                                                                      ByVal nSize As Long) As Long

Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long

Private Declare Function GetFileVersionInfoSize Lib "version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, _
                                                                                                   lpdwHandle As Long) As Long

Private Declare Function GetFileVersionInfo Lib "version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, _
                                                                                           ByVal dwHandle As Long, _
                                                                                           ByVal dwLen As Long, _
                                                                                           lpData As Any) As Long

Private Declare Function VerQueryValue Lib "version.dll" Alias "VerQueryValueA" (pBlock As Any, _
                                                                                 ByVal lpSubBlock As String, _
                                                                                 lplpBuffer As Any, _
                                                                                 puLen As Long) As Long

Private Declare Function lstrcpy Lib "kernel32" Alias "lstrcpyA" (ByVal lpString1 As String, _
                                                                  ByVal lpString2 As Long) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32" (Destination As Any, _
                                                  ByVal Source As Any, _
                                                  ByVal Length As Long)

Public Property Get FileName() As String

    FileName = mstrFileName

End Property

Public Property Let FileName(ByVal Value As String)

    mstrFileName = Value
    GetLanguageInfo

End Property

Public Property Get ModuleName() As String

    ModuleName = mstrModuleName

End Property

Public Property Let ModuleName(ByVal Value As String)


    mstrModuleName = Value
    GetLanguageInfo

End Property

Public Property Get Language() As Long

    If Len(mstrLangCP) = 8 Then
        Language = Val("&H" & Left$(mstrLangCP, 4))
    End If

End Property

Public Property Get CodePage() As Long

    If Len(mstrLangCP) = 8 Then
        CodePage = Val("&H" & Right$(mstrLangCP, 4))
    End If

End Property

Public Property Get CompanyName() As String

    CompanyName = GetValue("CompanyName")

End Property

Public Property Get FileDescription() As String

    FileDescription = GetValue("FileDescription")

End Property

Public Property Get FileVersion() As String

    FileVersion = GetValue("FileVersion")

End Property

Public Property Get InternalName() As String

    InternalName = GetValue("InternalName")

End Property

Public Property Get LegalCopyright() As String

    LegalCopyright = GetValue("LegalCopyright")

End Property

Public Property Get OriginalFileName() As String

    OriginalFileName = GetValue("OriginalFileName")

End Property

Public Property Get ProductName() As String

    ProductName = GetValue("ProductName")

End Property

Public Property Get ProductVersion() As String

    ProductVersion = GetValue("ProductVersion")

End Property

Public Property Get LegalTrademarks() As String

    LegalTrademarks = GetValue("LegalTrademarks")

End Property

Public Property Get LegalTrademarks1() As String

    LegalTrademarks1 = GetValue("LegalTrademarks1")

End Property

Public Property Get LegalTrademarks2() As String

    LegalTrademarks2 = GetValue("LegalTrademarks2")

End Property

Public Property Get Comments() As String

    Comments = GetValue("Comments")

End Property

Public Property Get PrivateBuild() As String

    PrivateBuild = GetValue("PrivateBuild")

End Property

Public Property Get SpecialBuild() As String

    SpecialBuild = GetValue("SpecialBuild")

End Property

Public Property Get FileFlagsRaw() As adhFileFlags

    FileFlagsRaw = mffi.dwFileFlags And mffi.dwFileFlagsMask

End Property

Public Property Get FileFlags() As String

Dim strOut  As String
Dim lngTemp As adhFileFlags

    lngTemp = mffi.dwFileFlags And mffi.dwFileFlagsMask

    If lngTemp And VS_FF_DEBUG Then
        strOut = strOut & "Debug "
    End If
    
    If lngTemp And VS_FF_PATCHED Then
        strOut = strOut & "Patched "
    End If
    
    If lngTemp And VS_FF_PRERELEASE Then
        strOut = strOut & "Pre-release "
    End If
    
    If lngTemp And VS_FF_PRIVATEBUILD Then
        strOut = strOut & "Private "
    End If
    
    If lngTemp And VS_FF_SPECIALBUILD Then
        strOut = strOut & "Special "
    End If

    FileFlags = strOut

End Property

Public Property Get FileOSRaw() As adhTargetOS

    FileOSRaw = mffi.dwFileOS

End Property

Public Property Get FileOS() As String

Dim lngTemp As Long
Dim strOut  As String

    lngTemp = mffi.dwFileOS

    If lngTemp And VOS__WINDOWS16 Then
        strOut = "Win16 API"
    ElseIf lngTemp And VOS__PM16 Then
        strOut = "PM16 API"
    ElseIf lngTemp And VOS__PM32 Then
        strOut = "PM32 API"
    ElseIf lngTemp And VOS__WINDOWS32 Then
        strOut = "Win32 API"
    Else
        strOut = "Unknown API"
    End If

    If lngTemp And VOS_DOS Then
        strOut = strOut & " on DOS"
    ElseIf lngTemp And VOS_OS216 Then
        strOut = strOut & " on OS2/16"
    ElseIf lngTemp And VOS_OS232 Then
        strOut = strOut & " on OS2/32"
    ElseIf lngTemp = VOS_NT Then
        strOut = strOut & " on Windows NT"
    End If

    FileOS = strOut

End Property

Public Property Get FileTypeRaw() As adhFileType

    FileTypeRaw = mffi.dwFileType

End Property

Public Property Get FileType() As String

Dim strOut As String

    Select Case mffi.dwFileType
    Case VFT_APP
        strOut = "Application"
    Case VFT_DLL
        strOut = "DLL"
    Case VFT_DRV
        strOut = "Driver"
    Case VFT_FONT
        strOut = "Font"
    Case VFT_STATIC_LIB
        strOut = "Static-link library"
    Case VFT_UNKNOWN
        strOut = "Unknown"
    Case VFT_VXD
        strOut = "VXD"
    End Select
    FileType = strOut

End Property

Public Property Get FileSubTypeRaw() As adhFileSubType

    FileSubTypeRaw = mffi.dwFileSubtype

End Property

Public Property Get FileSubType() As String

Dim strOut As String

    strOut = "Unknown"

    Select Case mffi.dwFileType
    Case VFT_APP
    Case VFT_DLL
    Case VFT_DRV
        Select Case mffi.dwFileSubtype
        Case VFT2_DRV_COMM
            strOut = "Communications"
        Case VFT2_DRV_DISPLAY
            strOut = "Display"
        Case VFT2_DRV_INPUTMETHOD
            strOut = "Input Method"
        Case VFT2_DRV_INSTALLABLE
            strOut = "Installable"
        Case VFT2_DRV_KEYBOARD
            strOut = "Keyboard"
        Case VFT2_DRV_LANGUAGE
            strOut = "Language"
        Case VFT2_DRV_MOUSE
            strOut = "Mouse"
        Case VFT2_DRV_NETWORK
            strOut = "Network"
        Case VFT2_DRV_PRINTER
            strOut = "Printer"
        Case VFT2_DRV_SOUND
            strOut = "Sound"
        Case VFT2_DRV_SYSTEM
            strOut = "System"
        Case VFT2_UNKNOWN
            strOut = "Unknown"
        End Select
    Case VFT_FONT
        Select Case mffi.dwFileSubtype
        Case VFT2_FONT_RASTER
            strOut = "Raster"
        Case VFT2_FONT_TRUETYPE
            strOut = "TrueType"
        Case VFT2_FONT_VECTOR
            strOut = "Vector"
        Case VFT2_UNKNOWN
            strOut = "Unknown"
        End Select
    Case VFT_STATIC_LIB
    Case VFT_UNKNOWN
    Case VFT_VXD
        strOut = mffi.dwFileSubtype
    End Select
    FileSubType = strOut

End Property

Private Function GetValue(ByVal strItem As String) As String

Dim strTemp       As String
Dim lngVerPointer As Long
Dim lngBufferLen  As Long
Dim strResult     As String

On Error GoTo Handler

    strTemp = "\StringFileInfo\" & mstrLangCP & "\" & strItem
    If VerQueryValue(mabytBuffer(0), strTemp, lngVerPointer, lngBufferLen) <> 0 Then
        strResult = Space$(lngBufferLen - 1)
        lstrcpy strResult, lngVerPointer
        strResult = adhTrimNull(strResult)
    End If

ExitHere:
    GetValue = strResult

Exit Function

Handler:
    strResult = vbNullString
    Resume ExitHere

End Function

Private Sub GetLanguageInfo()

Dim hInst         As Long
Dim lngValue      As Long
Dim lngBufferLen  As Long
Dim lngVerPointer As Long
Dim abytTemp()    As Byte

    If Len(mstrModuleName) > 0 Then
        hInst = GetModuleHandle(mstrModuleName)
        mstrFileName = Space$(260)
        If GetModuleFileName(hInst, mstrFileName, Len(mstrFileName)) = 0 Then
            Exit Sub
        End If
    End If

    If Len(mstrFileName) = 0 Then
        Exit Sub
    End If

    lngBufferLen = GetFileVersionInfoSize(mstrFileName, lngValue)
    If lngBufferLen = 0 Then
        Exit Sub
    End If

    ReDim mabytBuffer(0 To lngBufferLen - 1)
    If GetFileVersionInfo(mstrFileName, 0, lngBufferLen, mabytBuffer(0)) = 0 Then
        Exit Sub
    End If

    If VerQueryValue(mabytBuffer(0), "\VarFileInfo\Translation", lngVerPointer, lngBufferLen) = 0 Then
        Exit Sub
    End If

    ReDim abytTemp(0 To lngBufferLen - 1) As Byte
    RtlMoveMemory abytTemp(0), lngVerPointer, lngBufferLen
    mstrLangCP = ZeroPad(Hex$(abytTemp(1)), 2) & ZeroPad(Hex$(abytTemp(0)), 2) & ZeroPad(Hex$(abytTemp(3)), 2) & ZeroPad(Hex$(abytTemp(2)), 2)

    If mstrLangCP = "00000000" Then
        mstrLangCP = "040904E4"
    End If

    If VerQueryValue(mabytBuffer(0), "\", lngVerPointer, lngBufferLen) = 0 Then
        Exit Sub
    End If
    RtlMoveMemory mffi, lngVerPointer, lngBufferLen

End Sub

Private Function adhTrimNull(strVal As String) As String

Dim intPos As Integer

    intPos = InStr(1, strVal, vbNullChar)
    Select Case intPos
    Case Is > 1
        adhTrimNull = Left$(strVal, intPos - 1)
    Case 0
        adhTrimNull = strVal
    Case 1
        adhTrimNull = vbNullString
    End Select

End Function

Private Function ZeroPad(strValue As String, _
                         intLen As String) As String

    ZeroPad = Right$(String$(intLen, "0") & strValue, intLen)

End Function



