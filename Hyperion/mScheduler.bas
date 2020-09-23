Attribute VB_Name = "mScheduler"
Option Explicit

'/* api constants
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const SERVICE_ALL_ACCESS = SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS _
              Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE _
              Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL

'/* service state
Private Const SERVICE_STOPPED = &H1
Private Const SERVICE_START_PENDING = &H2
Private Const SERVICE_STOP_PENDING = &H3
Private Const SERVICE_RUNNING = &H4
Private Const SERVICE_CONTINUE_PENDING = &H5
Private Const SERVICE_PAUSE_PENDING = &H6
Private Const SERVICE_PAUSED = &H7
Private Const SERVICE_CONTROL_STOP = &H1
Private Const SERVICE_CONTROL_PAUSE = &H2
Private Const SERVICE_CONTROL_CONTINUE = &H3

    '/* service start type
Private Const SC_MANAGER_CONNECT = &H1
Private Const SERVICE_BOOT_START = &H0
Private Const SERVICE_SYSTEM_START = &H1
Private Const SERVICE_AUTO_START = &H2
Private Const SERVICE_DEMAND_START = &H3
Private Const SERVICE_DISABLED = &H4
Private Const SERVICE_NO_CHANGE = &HFFFFFFFF

'/* startup type
Private Enum ServiceStartType
    START_BOOT = &H0
    START_SYSTEM = &H1
    START_AUTO = &H2
    START_DEMAND = &H3
    START_DISABLED = &H4
End Enum

'/* status
Private Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type

Public Enum eDaysOfWeek
    Monday = 1
    Tuesday = 2
    Wednesday = 4
    Thursday = 8
    Friday = 16
    Saturday = 32
    Sunday = 64
End Enum

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal strMachineName As String, _
                                                                          ByVal strDBName As String, _
                                                                          ByVal lAccessReq As Long) As Long
                                                                          
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, _
                                                                      ByVal strServiceName As String, _
                                                                      ByVal lAccessReq As Long) As Long
                                                                      
Private Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, _
                                                                        ByVal lNumServiceArgs As Long, _
                                                                        ByVal strArgs As String) As Boolean
                                                                        
Private Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, _
                                                    ByVal lControlCode As Long, _
                                                    lpServiceStatus As SERVICE_STATUS) As Boolean
                                                    
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hHandle As Long) As Boolean

Private Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, _
                                                        lpServiceStatus As SERVICE_STATUS) As Boolean
                                                        
Private Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, _
                                                                                      ByVal dwServiceType As Long, _
                                                                                      ByVal dwStartType As ServiceStartType, _
                                                                                      ByVal dwErrorControl As Long, _
                                                                                      ByVal lpBinaryPathName As String, _
                                                                                      ByVal lpLoadOrderGroup As String, _
                                                                                      ByVal lpdwTagID As Long, _
                                                                                      ByVal lpDependencies As String, _
                                                                                      ByVal lpServiceStartName As String, _
                                                                                      ByVal lpPassword As String, _
                                                                                      ByVal lpDisplayName As String) As Boolean

Public Declare Sub Sleep Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Sub Schedule_Add(ByVal iInterval As Integer, _
                        Optional ByVal iDays As Integer)
'/* ain't wmi cool? no lame netapi here peeps..
'/* more info: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/wmisdk/wmi/wmi_start_page.asp
'/* for the uninitiated, start here: http://msdn.microsoft.com/library/default.asp?url=/library/en-us/dnclinic/html/scripting06112002.asp
'/* doesn't work on older platforms, so I provide an alternative method/dialog structure for 98/me (Revert_Mode)

Dim lResult         As Long
Dim objWMIService   As SWbemServices
Dim objNewJob       As SWbemObject

On Error Resume Next

    '/* test services, if they fail
    '/* throw switch in registry to
    '/* bypass on next run (with reset
    '/* switch in options?)
    With New clsLightning
        If Not .Read_DWord(HKEY_CURRENT_USER, m_sRegPath, "svfl") = 2 Then
            If Not Get_ServiceState("schedule") Then
                .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "svfl", 1
                Revert_Mode
                Exit Sub
            Else
                .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "svfl", 2
            End If
        Else
            Revert_Mode
            Exit Sub
        End If
    End With
    
    '/* create a new instance
    Set objWMIService = GetObject("winmgmts:{impersonationLevel=impersonate}!\\.\root\cimv2")
    '/* set service object
    Set objNewJob = objWMIService.Get("Win32_ScheduledJob")
    '/* vals: app/time - [offset utc]/repeating/days of week/interactive/jobid
    '*~wscript.echo:wof!
    
    Select Case iInterval
    Case 1 '/* every 12 hrs
        '/* 1am UCT
        lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
        True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , False, "117")
        '/* 1pm UCT
        lResult = objNewJob.Create(m_sAppPath & " -s", "********013000.000000-420", _
        True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , False, "117")
    
    Case 2 '/* every day: 1am UCT
        lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
        True, 1 Or 2 Or 4 Or 8 Or 16 Or 32 Or 64, , False, "117")
        
    Case 3 '/* user select - monday-sunday
        Select Case iDays
            Case 0
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 1, , False, "117")
            Case 1
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 2, , False, "117")
            Case 2
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 4, , False, "117")
            Case 3
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 8, , False, "117")
            Case 4
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 16, , False, "117")
            Case 5
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 32, , False, "117")
            Case 6
                lResult = objNewJob.Create(m_sAppPath & " -s", "********010000.000000-420", _
                True, 64, , False, "117")
            End Select
    End Select
    
    If Not lResult = 0 Then
        Revert_Mode
    End If

On Error GoTo 0

End Sub

Public Function Job_Exists() As Boolean
'/* test for success

Dim objTasks    As Object
Dim objTask     As Object

On Error Resume Next

    Set objTasks = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_ScheduledJob")
    For Each objTask In objTasks
        If LCase$(objTask.Command) = LCase$(m_sAppPath) Then
            Job_Exists = True
            Exit For
        End If
    Next

On Error GoTo 0

End Function

Public Sub Schedule_Remove()
'/* delete any scheduled tasks

Dim objTasks As Object
Dim objTask As Object
Dim strCmdLine As String

On Error Resume Next

    '/* connect
    Set objTasks = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_ScheduledJob")
    If Err Then
        Revert_Mode
        Exit Sub
    End If
    
    '/* loop through tasks
    For Each objTask In objTasks
        '/* on a match delete
        If LCase$(objTask.Command) = LCase$(m_sAppPath & " -s") Then
            Err.Clear
            objTask.Delete
            If Err Then
                Revert_Mode
                Exit Sub
            End If
        End If
    Next objTask

On Error GoTo 0

End Sub

Public Sub Revert_Mode()
'/* setup visual notifier
'/* and test update index

Dim mDate   As Date
Dim sDate   As String

On Error Resume Next

    '/* update option settings on forms
    With frmQuery
        .chkIndex.Caption = "Enable Index Reminder"
        .optIndex(0).Enabled = False
        .optIndex(1).Enabled = False
        .optIndex(2).Enabled = False
        .cbDays.Enabled = False
        .lblInfo(4).Visible = False
        .lblNotice.Caption = "Scheduler Not Supported on 98/ME"
        .lblInfo(3).Caption = "Your Drive has not been indexed. " & _
        "Click the 'Start Indexing' button to begin scanning the drive,  " & _
        "or select another drive from the list to index."
    End With
    
    With frmMain
        .chkIndex.Caption = "Enable Index Reminder"
        .optIndex(0).Enabled = False
        .optIndex(1).Enabled = False
        .optIndex(2).Enabled = False
        .cbDays.Enabled = False
        .lblInfo(0).Caption = "Scheduler Not Supported on 98/ME"
    End With
    
    '/* last reindex check
    With New clsLightning
        '/* get date index last updated
        sDate = .Read_String(HKEY_CURRENT_USER, m_sRegPath, "updchk")
        '/* if absent write new
        If Len(sDate) = 0 Then
            .Write_String HKEY_CURRENT_USER, m_sRegPath, "updchk", CStr(Now)
            Exit Sub
        Else
            '/* compare to todays date
            '/* if more then 7 days, arm auto updater
            If DateDiff(1, CDate(sDate), Now) > 7 Then
                .Write_DWord HKEY_CURRENT_USER, m_sRegPath, "schreset", 2
            End If
        End If
    End With

On Error GoTo 0

End Sub

Public Function Custom_Message(ByVal sCaption As String, _
                               ByVal sHeader As String, _
                               ByVal sMessage As String, _
                               ByVal fCaller As Form) As Integer
'/* reusable user notifications

Dim frm As New frmMessage

    With frm
        .Caption = sCaption
        '.Owner = frmMain
        .lblHeader.Caption = sHeader
        .lblBody.Caption = sMessage
        .Show vbModal, fCaller
    End With

    Custom_Message = m_iChoice

End Function

Private Function Start_Service(sService As String) As Boolean
'/* start a service

Dim lConHandle    As Long
Dim lSvcHandle    As Long

On Error GoTo Handler

    '/* get handle to service manager
    lConHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    '/* get service handle
    lSvcHandle = OpenService(lConHandle, sService, SERVICE_ALL_ACCESS)
    Start_Service = StartService(lSvcHandle, 0&, 0&)
    '/* cleanup
    CloseServiceHandle lSvcHandle
    CloseServiceHandle lConHandle

Handler:
On Error GoTo 0

End Function

Private Function Stop_Service(sService As String) As Boolean
'/* uhh.. stop a service

Dim lConHandle      As Long
Dim lSvcHandle      As Long
Dim svcStatus       As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lConHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lSvcHandle = OpenService(lConHandle, sService, SERVICE_ALL_ACCESS)
    '/* stop service
    Stop_Service = ControlService(lSvcHandle, SERVICE_CONTROL_STOP, svcStatus)
    '/* cleanup
    CloseServiceHandle lSvcHandle
    CloseServiceHandle lConHandle

Handler:
On Error GoTo 0

End Function

Private Function Resume_Service(sService As String) As Boolean
'/* resume paused

Dim lConHandle      As Long
Dim lSvcHandle      As Long
Dim svcStatus       As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lConHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lSvcHandle = OpenService(lConHandle, sService, SERVICE_ALL_ACCESS)
    '/* stop service
    Resume_Service = ControlService(lSvcHandle, SERVICE_CONTROL_CONTINUE, svcStatus)
    '/* cleanup
    CloseServiceHandle lSvcHandle
    CloseServiceHandle lConHandle

Handler:
On Error GoTo 0

End Function

Private Function Query_Service(sService As String) As Long
'/* get service state

Dim lConHandle      As Long
Dim lSvcHandle      As Long
Dim svcStatus       As SERVICE_STATUS

On Error GoTo Handler

    '/* get app and service handles
    lConHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lSvcHandle = OpenService(lConHandle, sService, SERVICE_QUERY_STATUS)
    '/* query service status
    QueryServiceStatus lSvcHandle, svcStatus
    '/* cleanup
    CloseServiceHandle lSvcHandle
    CloseServiceHandle lConHandle
    '/* return service state
    '/* 0 - not exist, 1 - stopped, 2 - paused, 3 - waiting, 4 - running
    Query_Service = svcStatus.dwCurrentState

Handler:
On Error GoTo 0

End Function

Private Function Set_ServiceType(ByVal sService As String, _
                                  ByVal lType As Long) As Boolean

'/* set params
Dim lConHandle    As Long
Dim lSvcHandle   As Long

On Error GoTo Handler

    '/* get app and service handles
    lConHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    lSvcHandle = OpenService(lConHandle, sService, SERVICE_CHANGE_CONFIG)
    '/* change startup type
    Set_ServiceType = ChangeServiceConfig(lSvcHandle, SERVICE_NO_CHANGE, lType, _
    SERVICE_NO_CHANGE, vbNullString, vbNullString, 0&, vbNullString, vbNullString, vbNullString, vbNullString)
    '/* cleanup
    CloseServiceHandle lSvcHandle
    CloseServiceHandle lConHandle

Handler:
On Error GoTo 0

End Function

Public Function Get_ServiceState(ByVal sSrvName As String) As Boolean
'/* query service with wait timer loop

Dim lResult     As Long
Dim iFail       As Integer

On Error GoTo Handler

        lResult = Query_Service(sSrvName)
        Select Case lResult
        Case 0
        '/* does not exist?
            GoTo Handler
        Case 1
        '/* stopped - start and set to automatic
            If Not Start_Service(sSrvName) And _
            Set_ServiceType(sSrvName, &H2) Then GoTo Handler
                    
        Case 2
            '/* waiting to start
            Do
                iFail = iFail + 1
                Sleep 200
                DoEvents
                '/* if it takes more then 2 seconds to start then
                '/* it may be a dependancy issue  or other unknown
                '/* so best to bail
                If Query_Service(sSrvName) Then
                    Get_ServiceState = True
                    Exit Do
                End If
            Loop Until iFail = 10
            
        Case 3
        '/* waiting to stop - start and set to automatic
            If Not Start_Service(sSrvName) And _
            Set_ServiceType(sSrvName, &H2) Then GoTo Handler
            
        Case 4
            '/* running
            Get_ServiceState = True
                    
        Case 7
            '/* paused - resume and set to automatic
            If Not Resume_Service(sSrvName) And _
            Set_ServiceType(sSrvName, &H2) Then GoTo Handler
        End Select
        
Exit Function

Handler:
On Error GoTo 0

End Function

