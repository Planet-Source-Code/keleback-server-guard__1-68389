VERSION 5.00
Begin VB.Form FrmServices 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Kayýtlý Hizmetler"
   ClientHeight    =   4680
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   6105
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4680
   ScaleWidth      =   6105
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton BtnCancel 
      Caption         =   "Vazgeç"
      Height          =   375
      Left            =   1320
      TabIndex        =   2
      Top             =   4200
      Width           =   1095
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "Tamam"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   4200
      Width           =   1095
   End
   Begin VB.ListBox List1 
      Height          =   3960
      Left            =   120
      MultiSelect     =   1  'Simple
      TabIndex        =   0
      Top             =   120
      Width           =   5895
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Çift týklayarak istediðiniz hizmetin þu anki durumu hakkýnda bilgi alabilirsiniz."
      BeginProperty Font 
         Name            =   "Small Fonts"
         Size            =   6.75
         Charset         =   162
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   4200
      Visible         =   0   'False
      Width           =   3375
   End
   Begin VB.Menu MnuService 
      Caption         =   "Hizmet"
      Visible         =   0   'False
      Begin VB.Menu MnuQuery 
         Caption         =   "Sorgula"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuStart 
         Caption         =   "Baþlat"
      End
      Begin VB.Menu MnuStop 
         Caption         =   "Durdur"
      End
      Begin VB.Menu MnuPause 
         Caption         =   "Duraklat"
      End
      Begin VB.Menu MnuResume 
         Caption         =   "Sürdür"
      End
   End
End
Attribute VB_Name = "FrmServices"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Const ERROR_MORE_DATA = 234
Const SERVICE_ACTIVE = &H1
Const SERVICE_INACTIVE = &H2
Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Const SERVICE_WIN32_OWN_PROCESS As Long = &H10
Const SERVICE_WIN32_SHARE_PROCESS As Long = &H20
Const SERVICE_WIN32 As Long = SERVICE_WIN32_OWN_PROCESS + SERVICE_WIN32_SHARE_PROCESS
Private Type SERVICE_STATUS
    dwServiceType               As Long
    dwCurrentState              As Long
    dwControlsAccepted          As Long
    dwWin32ExitCode             As Long
    dwServiceSpecificExitCode   As Long
    dwCheckPoint                As Long
    dwWaitHint                  As Long
End Type
Private Type ENUM_SERVICE_STATUS
    lpServiceName               As Long
    lpDisplayName               As Long
    ServiceStatus               As SERVICE_STATUS
End Type
Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function EnumServicesStatus Lib "advapi32.dll" Alias "EnumServicesStatusA" (ByVal hSCManager As Long, ByVal dwServiceType As Long, ByVal dwServiceState As Long, lpServices As Any, ByVal cbBufSize As Long, pcbBytesNeeded As Long, lpServicesReturned As Long, lpResumeHandle As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (szDest As String, szcSource As Long) As Long

Dim ArrRegSvc()                 As String
Dim MousePressed                As Boolean
Dim ObjService                  As ClsService
Private Sub BtnCancel_Click()

    Unload Me
    
End Sub

Private Sub BtnOK_Click()

    Dim FlgExist    As Boolean
    Dim AddSvcCnt   As Integer
    
    For AddSvcCnt = 0 To List1.ListCount - 1
        If List1.Selected(AddSvcCnt) Then
            If FrmSetup.LstServices.ListCount > 0 Then
                For X = 0 To UBound(ArrServices, 2) - 1
                    If StrComp(ArrRegSvc(List1.ListIndex), ArrServices(0, X)) = 0 Then
                        FlgExist = True
                        Exit For
                    End If
                Next X
            End If
            If Not FlgExist Then
                ReDim Preserve ArrServices(1, FrmSetup.LstServices.ListCount)
                ArrServices(0, FrmSetup.LstServices.ListCount) = ArrRegSvc(AddSvcCnt)
                ArrServices(1, FrmSetup.LstServices.ListCount) = List1.List(AddSvcCnt)
                FrmSetup.LstServices.AddItem List1.List(AddSvcCnt)
            End If
        End If
        FlgExist = False
    Next AddSvcCnt
    Unload Me
    
End Sub

Private Sub Form_Load()

    Dim hSCM As Long, lpEnumServiceStatus() As ENUM_SERVICE_STATUS, lngServiceStatusInfoBuffer As Long
    Dim strServiceName As String * 250, lngBytesNeeded As Long, lngServicesReturned As Long
    Dim hNextUnreadEntry As Long, lngStructsNeeded As Long, lngResult As Long, i As Long
    
    Set ObjService = New ClsService
    
    hSCM = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_ENUMERATE_SERVICE)
    If hSCM = 0 Then
        MsgBox "OpenSCManager failed. LastDllError = " & CStr(Err.LastDllError)
        Exit Sub
    End If

    hNextUnreadEntry = 0
    lngResult = EnumServicesStatus(hSCM, SERVICE_WIN32, SERVICE_ACTIVE Or SERVICE_INACTIVE, ByVal &H0, &H0, lngBytesNeeded, lngServicesReturned, hNextUnreadEntry)

    If Not Err.LastDllError = ERROR_MORE_DATA Then
        MsgBox "LastDLLError = " & CStr(Err.LastDllError)
        Exit Sub
    End If

    lngStructsNeeded = lngBytesNeeded / Len(lpEnumServiceStatus(0)) + 1

    ReDim lpEnumServiceStatus(lngStructsNeeded - 1)
    lngServiceStatusInfoBuffer = lngStructsNeeded * Len(lpEnumServiceStatus(0))
    hNextUnreadEntry = 0
    lngResult = EnumServicesStatus(hSCM, SERVICE_WIN32, SERVICE_ACTIVE Or SERVICE_INACTIVE, lpEnumServiceStatus(0), lngServiceStatusInfoBuffer, lngBytesNeeded, lngServicesReturned, hNextUnreadEntry)
    If lngResult = 0 Then
        MsgBox "EnumServicesStatus failed. LastDllError = " & CStr(Err.LastDllError)
        Exit Sub
    End If

    ReDim ArrRegSvc(lngServicesReturned - 1)
    For i = 0 To lngServicesReturned - 1
        lngResult = lstrcpy(ByVal strServiceName, ByVal lpEnumServiceStatus(i).lpDisplayName)
        List1.AddItem StripTerminator(strServiceName)
        lngResult = lstrcpy(ByVal strServiceName, ByVal lpEnumServiceStatus(i).lpServiceName)
        ArrRegSvc(i) = StripTerminator(strServiceName)
    Next i
    CloseServiceHandle (hSCM)
End Sub
Function StripTerminator(sInput As String) As String
    
    Dim ZeroPos As Integer
    ZeroPos = InStr(1, sInput, Chr$(0))
    If ZeroPos > 0 Then
        StripTerminator = Left$(sInput, ZeroPos - 1)
    Else
        StripTerminator = sInput
    End If
    
End Function


Private Sub List1_DblClick()

    ObjService.ServiceName = ArrRegSvc(List1.ListIndex)
    MsgBox ObjService.QueryService
    
End Sub


Private Sub List1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MousePressed = True
    
End Sub


Private Sub List1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MousePressed Then
        MousePressed = False
        If Button = 2 Then Me.PopupMenu MnuService, , , , MnuQuery
    End If
    
End Sub


Private Sub MnuPause_Click()
    
    ObjService.ServiceName = ArrRegSvc(List1.ListIndex)
    ObjService.PauseService
    
End Sub

Private Sub MnuQuery_Click()

    ObjService.ServiceName = ArrRegSvc(List1.ListIndex)
    MsgBox ObjService.QueryService
    
End Sub


Private Sub MnuResume_Click()
    
    ObjService.ServiceName = ArrRegSvc(List1.ListIndex)
    ObjService.ResumeService
    
End Sub

Private Sub MnuStart_Click()
    
    ObjService.ServiceName = ArrRegSvc(List1.ListIndex)
    ObjService.StartService
    
End Sub


Private Sub MnuStop_Click()
    
    ObjService.ServiceName = ArrRegSvc(List1.ListIndex)
    ObjService.StopService
    
End Sub


