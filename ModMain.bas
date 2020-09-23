Attribute VB_Name = "ModMain"
Public Const NIM_ADD = &H0
Public Const NIM_MODIFY = &H1
Public Const NIM_DELETE = &H2
Public Const NIF_MESSAGE = &H1
Public Const NIF_ICON = &H2
Public Const NIF_TIP = &H4

Public Const WM_LBUTTONDBLCLK = &H203
Public Const WM_LBUTTONDOWN = &H201
Public Const WM_RBUTTONUP = &H205

' registry constants
Public Const READ_CONTROL = &H20000
Public Const KEY_SET_VALUE = &H2
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const SYNCHRONIZE = &H100000
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_READ = ((STANDARD_RIGHTS_READ Or KEY_QUERY_VALUE Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY) And (Not SYNCHRONIZE))
Public Const HKEY_CURRENT_USER = &H80000001
Public Const REG_SZ = 1
Public Const SND_FILENAME = &H20000
Public Const SND_ASYNC = &H1
Public Const SND_MEMORY = &H4
Public Const SND_NODEFAULT = &H2

Public Declare Function Shell_NotifyIcon Lib "shell32" Alias "Shell_NotifyIconA" (ByVal dwMessage As Long, pnid As NOTIFYICONDATA) As Boolean
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, ByVal lpSecurityAttributes As Long, phkResult As Long, lpdwDisposition As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long
    
Type SvcRecord
    ServiceName         As String * 50
    ServiceDispName     As String * 100
End Type

Type SchRecord
    Time                As String * 5
End Type

Public Type NOTIFYICONDATA
    cbSize              As Long
    hWnd                As Long
    uId                 As Long
    uFlags              As Long
    ucallbackMessage    As Long
    hIcon               As Long
    szTip               As String * 64
End Type
    
Public ArrServices()    As String               '2 Dimensional (Service Name, Display Name)
Public ArrSchedule()    As String               '1 Dimensional
Public TrayIcon         As NOTIFYICONDATA
Public RecHolder        As SvcRecord
Public SchHolder        As SchRecord
Public Function AutoStartDelete() As Long
    
    Dim hKey As Long
    Dim lRtn As Long
    
    lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
    If lRtn = False Then AutoStartDelete = RegDeleteValue(hKey, App.EXEName)
    
End Function
Public Function AutoStartAdd() As Long

    Dim hKey As Long
    Dim lRtn As Long
    Dim sPathApp As String
    
    lRtn = RegCreateKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
        ByVal 0&, ByVal 0&, ByVal 0&, KEY_WRITE, ByVal 0&, hKey, ByVal 0&)
     
    If lRtn = False Then
        sPathApp = Chr(34) & App.Path & Chr(92) & App.EXEName & ".exe" & Chr(34)
        lRtn = RegSetValueEx(hKey, App.EXEName, 0, REG_SZ, ByVal sPathApp, Len(sPathApp))
    End If
    
End Function
Public Function IsAutoStart() As Boolean

    Dim hKey As Long
    Dim lType As Long
    Dim sValue As String
    
    sValue = App.EXEName
    If RegOpenKeyEx(HKEY_CURRENT_USER, "Software\Microsoft\Windows\CurrentVersion\Run", _
        0, KEY_READ, hKey) = False Then
        If RegQueryValueEx(hKey, sValue, ByVal 0&, lType, ByVal 0&, ByVal 0&) = False Then
            IsAutoStart = True
            RegCloseKey hKey
        End If
    End If
    
End Function
Public Sub TrayIconCreate(FrmHost As Form, Optional Caption As String)
    
    If Caption = "" Then Caption = App.ProductName
    With TrayIcon
        .cbSize = Len(TrayIcon)
        .hWnd = FrmHost.hWnd
        .uId = 1&
        .uFlags = NIF_ICON Or NIF_TIP Or NIF_MESSAGE
        .ucallbackMessage = WM_LBUTTONDOWN
        .hIcon = FrmHost.Icon
        .szTip = Caption & Chr$(0)
    End With
    Shell_NotifyIcon NIM_ADD, TrayIcon
    FrmHost.Hide
    
End Sub


