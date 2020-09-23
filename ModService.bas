Attribute VB_Name = "modService"
Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal strMachineName As String, ByVal strDBName As String, ByVal lAccessReq As Long) As Long
Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal strServiceName As String, ByVal lAccessReq As Long) As Long
Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal lNumServiceArgs As Long, ByVal strArgs As String) As Boolean
Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal lControlCode As Long, lpServiceStatus As SERVICE_STATUS) As Boolean
Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hHandle As Long) As Boolean
Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Boolean
Declare Function ChangeServiceConfig Lib "advapi32.dll" Alias "ChangeServiceConfigA" (ByVal hService As Long, ByVal dwServiceType As Long, ByVal dwStartType As ServiceStartType, ByVal dwErrorControl As Long, ByVal lpBinaryPathName As String, ByVal lpLoadOrderGroup As String, ByVal lpdwTagID As Long, ByVal lpDependencies As String, ByVal lpServiceStartName As String, ByVal lpPassword As String, ByVal lpDisplayName As String) As Boolean

Type SERVICE_STATUS
    dwServiceType               As Long
    dwCurrentState              As Long
    dwControlsAccepted          As Long
    dwWin32ExitCode             As Long
    dwServiceSpecificExitCode   As Long
    dwCheckPoint                As Long
    dwWaitHint                  As Long
End Type

Const SERVICE_QUERY_CONFIG = &H1
Const SERVICE_CHANGE_CONFIG = &H2
Const SERVICE_QUERY_STATUS = &H4
Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Const SERVICE_START = &H10
Const SERVICE_STOP = &H20
Const SERVICE_PAUSE_CONTINUE = &H40
Const SERVICE_INTERROGATE = &H80
Const SERVICE_USER_DEFINED_CONTROL = &H100
Const SERVICE_ALL_ACCESS = SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS _
    Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE _
    Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL
Const SERVICE_STOPPED = &H1
Const SERVICE_START_PENDING = &H2
Const SERVICE_STOP_PENDING = &H3
Const SERVICE_RUNNING = &H4
Const SERVICE_CONTINUE_PENDING = &H5
Const SERVICE_PAUSE_PENDING = &H6
Const SERVICE_PAUSED = &H7
Const SERVICE_CONTROL_STOP = &H1
Const SERVICE_CONTROL_PAUSE = &H2
Const SERVICE_CONTROL_CONTINUE = &H3
Const SC_MANAGER_CONNECT = &H1
Const SERVICE_BOOT_START = &H0
Const SERVICE_SYSTEM_START = &H1
Const SERVICE_AUTO_START = &H2
Const SERVICE_DEMAND_START = &H3
Const SERVICE_DISABLED = &H4
Const SERVICE_NO_CHANGE = &HFFFFFFFF

Public Function StartSvc(strServiceName As String) As Boolean

    Dim scHandle As Long, svcHandle As Long
    Dim RetCode As Long
    
    scHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    svcHandle = OpenService(scHandle, strServiceName, SERVICE_ALL_ACCESS)
    StartSvc = StartService(svcHandle, 0&, 0&)
    RetCode = CloseServiceHandle(svcHandle)
    RetCode = CloseServiceHandle(scHandle)

End Function

Public Function StopSvc(strServiceName As String) As Boolean

    Dim scHandle As Long, svcHandle As Long, svcStatus As SERVICE_STATUS
    Dim RetCode As Long
    
    scHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    svcHandle = OpenService(scHandle, strServiceName, SERVICE_ALL_ACCESS)
    StopSvc = ControlService(svcHandle, SERVICE_CONTROL_STOP, svcStatus)
    RetCode = CloseServiceHandle(svcHandle)
    RetCode = CloseServiceHandle(scHandle)

End Function

Public Function PauseSvc(strServiceName As String) As Boolean

    Dim scHandle As Long, svcHandle As Long, svcStatus As SERVICE_STATUS
    Dim RetCode As Long
    
    scHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    svcHandle = OpenService(scHandle, strServiceName, SERVICE_ALL_ACCESS)
    PauseSvc = ControlService(svcHandle, SERVICE_CONTROL_PAUSE, svcStatus)
    RetCode = CloseServiceHandle(svcHandle)
    RetCode = CloseServiceHandle(scHandle)

End Function

Public Function ResumeSvc(strServiceName As String) As Boolean

    Dim scHandle As Long, svcHandle As Long, svcStatus As SERVICE_STATUS
    Dim RetCode As Long
        
    scHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    svcHandle = OpenService(scHandle, strServiceName, SERVICE_ALL_ACCESS)
    ResumeSvc = ControlService(svcHandle, SERVICE_CONTROL_CONTINUE, svcStatus)
    RetCode = CloseServiceHandle(svcHandle)
    RetCode = CloseServiceHandle(scHandle)

End Function

Public Function QuerySvc(strServiceName As String) As Long

    Dim scHandle As Long, svcHandle As Long, svcStatus As SERVICE_STATUS
    Dim RetCode As Long
    
    scHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    svcHandle = OpenService(scHandle, strServiceName, SERVICE_QUERY_STATUS)
    RetCode = QueryServiceStatus(svcHandle, svcStatus)
    RetCode = CloseServiceHandle(svcHandle)
    RetCode = CloseServiceHandle(scHandle)
    QuerySvc = svcStatus.dwCurrentState

End Function

Public Function SetSvcStartType(strServiceName As String, svcStartType As ServiceStartType) As Boolean

    Dim RetCode As Long
    
    scHandle = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CONNECT)
    svcHandle = OpenService(scHandle, strServiceName, SERVICE_CHANGE_CONFIG)
    SetSvcStartType = ChangeServiceConfig(svcHandle, SERVICE_NO_CHANGE, svcStartType, SERVICE_NO_CHANGE, vbNullString, vbNullString, 0&, vbNullString, vbNullString, vbNullString, vbNullString)
    RetCode = CloseServiceHandle(svcHandle)
    RetCode = CloseServiceHandle(scHandle)

End Function

Public Function EnumSvcState(svcStatus As Long) As String

Select Case svcStatus
    Case SERVICE_STOPPED
        EnumSvcState = "Stopped"
    Case SERVICE_START_PENDING
        EnumSvcState = "Starting"
    Case SERVICE_STOP_PENDING
        EnumSvcState = "Stopping"
    Case SERVICE_RUNNING
        EnumSvcState = "Started"
    Case SERVICE_CONTINUE_PENDING
        EnumSvcState = "Continuing"
    Case SERVICE_PAUSE_PENDING
        EnumSvcState = "Pausing"
    Case SERVICE_PAUSED
        EnumSvcState = "Paused"
    Case Else
        EnumSvcState = "Query Not Successful"
End Select

End Function

