Attribute VB_Name = "mStatistics"
Option Explicit

Private m_dFspace     As Double
Private m_dTspace     As Double
Private m_dUspace     As Double

Private Type ULong
    Byte1               As Byte
    Byte2               As Byte
    Byte3               As Byte
    Byte4               As Byte
End Type

Private Type LargeInt
    LoDWord             As ULong
    HiDWord             As ULong
    LoDWord2            As ULong
    HiDWord2            As ULong
End Type

Private Declare Function GetDiskFreeSpaceEx Lib "kernel32" Alias "GetDiskFreeSpaceExA" (ByVal lpRootPathName As String, _
                                                                                        FreeBytesAvailableToCaller As LargeInt, _
                                                                                        TotalNumberOfBytes As LargeInt, _
                                                                                        TotalNumberOfFreeBytes As LargeInt) As Long


Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                                                               ByVal lpOperation As String, _
                                                                               ByVal lpFile As String, _
                                                                               ByVal lpParameters As String, _
                                                                               ByVal lpDirectory As String, _
                                                                               ByVal nShowCmd As Long) As Long

Private Function CULong(Byte1 As Byte, _
                        Byte2 As Byte, _
                        Byte3 As Byte, _
                        Byte4 As Byte) As Double
'/* long number type for calculation

    CULong = Byte4 * 2 ^ 24 + Byte3 * 2 ^ 16 + Byte2 * 2 ^ 8 + Byte1

End Function

Private Function GetDiskSpace(ByVal sPath As String) As Long
'/* get drive stats

Dim nFreeBytesToCaller As LargeInt
Dim nTotalBytes        As LargeInt
Dim nTotalFreeBytes    As LargeInt

On Error Resume Next

    GetDiskFreeSpaceEx sPath, nFreeBytesToCaller, nTotalBytes, nTotalFreeBytes
    m_dFspace = CULong(nFreeBytesToCaller.HiDWord.Byte1, nFreeBytesToCaller.HiDWord.Byte2, nFreeBytesToCaller.HiDWord.Byte3, nFreeBytesToCaller.HiDWord.Byte4) * 2 ^ 32 + CULong(nFreeBytesToCaller.LoDWord.Byte1, nFreeBytesToCaller.LoDWord.Byte2, nFreeBytesToCaller.LoDWord.Byte3, nFreeBytesToCaller.LoDWord.Byte4)
    m_dTspace = CULong(nTotalBytes.HiDWord.Byte1, nTotalBytes.HiDWord.Byte2, nTotalBytes.HiDWord.Byte3, nTotalBytes.HiDWord.Byte4) * 2 ^ 32 + CULong(nTotalBytes.LoDWord.Byte1, nTotalBytes.LoDWord.Byte2, nTotalBytes.LoDWord.Byte3, nTotalBytes.LoDWord.Byte4)
    m_dUspace = (m_dTspace - m_dFspace)
    m_dUspace = ((m_dUspace / 1024) / 1024)
    m_dFspace = ((m_dFspace / 1024) / 1024)
    GetDiskSpace = CLng(m_dUspace)

On Error GoTo 0

End Function

Public Function Drive_Used(ByVal sPath As String) As Long
'/* pass value back to caller

Dim lTInterval As Long

    lTInterval = GetDiskSpace(sPath)
    If lTInterval = 0 Then
        lTInterval = 1000
    End If
    Drive_Used = lTInterval

End Function



