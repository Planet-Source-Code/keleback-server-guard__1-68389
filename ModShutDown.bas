Attribute VB_Name = "ModShutDown"
' Shutdown Flags
Const EWX_LOGOFF = 0
Const EWX_SHUTDOWN = 1
Const EWX_REBOOT = 2
Const EWX_FORCE = 4
Const SE_PRIVILEGE_ENABLED = &H2
Const TokenPrivileges = 3
Const TOKEN_ASSIGN_PRIMARY = &H1
Const TOKEN_DUPLICATE = &H2
Const TOKEN_IMPERSONATE = &H4
Const TOKEN_QUERY = &H8
Const TOKEN_QUERY_SOURCE = &H10
Const TOKEN_ADJUST_PRIVILEGES = &H20
Const TOKEN_ADJUST_GROUPS = &H40
Const TOKEN_ADJUST_DEFAULT = &H80
Const SE_SHUTDOWN_NAME = "SeShutdownPrivilege"
Const ANYSIZE_ARRAY = 1
Private Type LARGE_INTEGER
    lowpart                     As Long
    highpart                    As Long
End Type
Private Type Luid
    lowpart                     As Long
    highpart                    As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    'pLuid As Luid
    pLuid                       As LARGE_INTEGER
    Attributes                  As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount              As Long
    Privileges(ANYSIZE_ARRAY)   As LUID_AND_ATTRIBUTES
End Type
Private Declare Function InitiateSystemShutdown Lib "advapi32.dll" Alias "InitiateSystemShutdownA" (ByVal lpMachineName As String, ByVal lpMessage As String, ByVal dwTimeout As Long, ByVal bForceAppsClosed As Long, ByVal bRebootAfterShutdown As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As String, ByVal lpName As String, lpLuid As LARGE_INTEGER) As Long
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPrivileges As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Private Declare Function GetLastError Lib "kernel32" () As Long
Public Function InitiateShutdownMachine(ByVal Machine As String, Optional Force As Variant, Optional Restart As Variant, Optional AllowLocalShutdown As Variant, Optional Delay As Variant, Optional Message As Variant) As Boolean
    
    Dim hProc As Long
    Dim OldTokenStuff As TOKEN_PRIVILEGES
    Dim OldTokenStuffLen As Long
    Dim NewTokenStuff As TOKEN_PRIVILEGES
    Dim NewTokenStuffLen As Long
    Dim pSize As Long
    If IsMissing(Force) Then Force = False
    If IsMissing(Restart) Then Restart = True
    If IsMissing(AllowLocalShutdown) Then AllowLocalShutdown = False
    If IsMissing(Delay) Then Delay = 0
    If IsMissing(Message) Then Message = ""

    If InStr(Machine, "\\") = 1 Then
        Machine = Right(Machine, Len(Machine) - 2)
    End If

    If (LCase(GetMyMachineName) = LCase(Machine)) Then
        If AllowLocalShutdown = False Then Exit Function
        If OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, hProc) = 0 Then
            MsgBox "OpenProcessToken Error: " & GetLastError()
            Exit Function
        End If
        If LookupPrivilegeValue(vbNullString, SE_SHUTDOWN_NAME, OldTokenStuff.Privileges(0).pLuid) = 0 Then
            MsgBox "LookupPrivilegeValue Error: " & GetLastError()
            Exit Function
        End If
        NewTokenStuff = OldTokenStuff
        NewTokenStuff.PrivilegeCount = 1
        NewTokenStuff.Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        NewTokenStuffLen = Len(NewTokenStuff)
        pSize = Len(NewTokenStuff)
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, NewTokenStuffLen, OldTokenStuff, OldTokenStuffLen) = 0 Then
            MsgBox "AdjustTokenPrivileges Error: " & GetLastError()
            Exit Function
        End If
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
        NewTokenStuff.Privileges(0).Attributes = 0
        If AdjustTokenPrivileges(hProc, False, NewTokenStuff, Len(NewTokenStuff), OldTokenStuff, Len(OldTokenStuff)) = 0 Then
            Exit Function
        End If
    Else
        If InitiateSystemShutdown("\\" & Machine, Message, Delay, Force, Restart) = 0 Then
            Exit Function
        End If
    End If
    InitiateShutdownMachine = True
End Function
Function GetMyMachineName() As String
    
    Dim sLen As Long
    
    GetMyMachineName = Space(100)
    sLen = 100
    If GetComputerName(GetMyMachineName, sLen) Then
        GetMyMachineName = Left(GetMyMachineName, sLen)
    End If
    
End Function
