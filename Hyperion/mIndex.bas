Attribute VB_Name = "mIndex"
Option Explicit

Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, _
                                                                                               ByVal lpBuffer As String) As Long

Public Function Scan_Time(ByVal sDrive As String, _
                          ByRef sUsed As String, _
                          ByRef sScnTm As String)
'/* good use of byref

Dim lEstimate   As Long
Dim lRefNum     As Long

    lRefNum = Drive_Used(sDrive)
    lEstimate = lRefNum / 190
    sUsed = "Used Space: " & lRefNum & " MegaBytes"
    sScnTm = "Scan Time Est: " & lEstimate & " Seconds to Index"
    
End Function

Public Function File_Exists(ByVal sPath As String) As Boolean
'/* test file path

    If Len(Dir(sPath)) > 0 Then
        File_Exists = True
    End If
    
End Function

Public Function Get_Drive(ByVal iDrive As Long) As String
'/* get drive letter from index

Dim sDrive As String
Dim lBuffer As Long

    '/* get the buffer size
    lBuffer = GetLogicalDriveStrings(0, sDrive)
    '/* set string len
    sDrive = String$(lBuffer, 0)
    '/* get the drive list
    GetLogicalDriveStrings lBuffer, sDrive
    '/* a: drive
    If iDrive = 0 Then
        Get_Drive = left$(sDrive, 3)
    Else
        '/* format and return
        iDrive = iDrive * 4
        Get_Drive = Mid$(sDrive, InStr((iDrive), sDrive, vbNullChar) + 1, 3)
    End If
    
End Function

Public Sub Unload_All()
'/* unload all forms and end

Dim frm As Form

On Error GoTo Handler

    For Each frm In Forms
        Unload frm
    Next

Handler:
End

End Sub


