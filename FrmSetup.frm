VERSION 5.00
Begin VB.Form FrmSetup 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Ayarlar"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6870
   Icon            =   "FrmSetup.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3015
   ScaleWidth      =   6870
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox ChkRestart 
      Caption         =   "Yeniden baþlat"
      Height          =   195
      Left            =   1200
      TabIndex        =   15
      Top             =   2670
      Width           =   1455
   End
   Begin VB.ComboBox CmbResHr 
      Height          =   315
      Left            =   2760
      Style           =   2  'Dropdown List
      TabIndex        =   14
      Top             =   2610
      Width           =   615
   End
   Begin VB.ComboBox CmbResMin 
      Height          =   315
      Left            =   3480
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2610
      Width           =   615
   End
   Begin VB.CheckBox ChkAutoStrt 
      Caption         =   "Windows baþladýðýnda baþla"
      Height          =   195
      Left            =   4440
      TabIndex        =   12
      Top             =   2670
      Width           =   2415
   End
   Begin VB.Timer Timer1 
      Interval        =   60000
      Left            =   -240
      Top             =   2280
   End
   Begin VB.CommandButton BtnOK 
      Caption         =   "Tamam"
      Height          =   375
      Left            =   60
      TabIndex        =   11
      Top             =   2580
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Ýzlenecek Hizmetler"
      Height          =   2415
      Left            =   2940
      TabIndex        =   9
      Top             =   60
      Width           =   3855
      Begin VB.CommandButton BtnRemove 
         Caption         =   "Çýkar"
         Height          =   255
         Left            =   960
         TabIndex        =   6
         Top             =   2040
         Width           =   735
      End
      Begin VB.CommandButton BtnAdd 
         Caption         =   "Ekle"
         Height          =   255
         Left            =   120
         TabIndex        =   7
         Top             =   2040
         Width           =   735
      End
      Begin VB.ListBox LstServices 
         Height          =   1620
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   3615
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Çizelge"
      Height          =   2415
      Left            =   60
      TabIndex        =   8
      Top             =   60
      Width           =   2775
      Begin VB.ComboBox CmbMin 
         Height          =   315
         ItemData        =   "FrmSetup.frx":15162
         Left            =   2040
         List            =   "FrmSetup.frx":15164
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   2010
         Width           =   615
      End
      Begin VB.ComboBox CmbHour 
         Height          =   315
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   2010
         Width           =   615
      End
      Begin VB.CommandButton BtnAddSchedule 
         Caption         =   "Ekle"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   2040
         Width           =   510
      End
      Begin VB.CommandButton BtnRemoveSchedule 
         Caption         =   "Çýkar"
         Height          =   255
         Left            =   720
         TabIndex        =   1
         Top             =   2040
         Width           =   510
      End
      Begin VB.ListBox LstSchedule 
         Height          =   1620
         Left            =   120
         Sorted          =   -1  'True
         TabIndex        =   0
         Top             =   240
         Width           =   2535
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   ":"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   162
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   1950
         TabIndex        =   10
         Top             =   2047
         Width           =   75
      End
   End
   Begin VB.Menu MnuMain 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu MnuSetup 
         Caption         =   "Ayarlar"
      End
      Begin VB.Menu MnuServices 
         Caption         =   "Hizmetler"
      End
      Begin VB.Menu MnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu MnuAbout 
         Caption         =   "Hakkýnda"
      End
      Begin VB.Menu MnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Çýkýþ"
      End
   End
End
Attribute VB_Name = "FrmSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim SvcObj  As ClsService
Public Sub CheckServices()
    
    Dim SvcIdx As Integer
    
    For SvcIdx = 0 To LstServices.ListCount - 1
        SvcObj.ServiceName = ArrServices(0, SvcIdx)
        If SvcObj.QueryService <> "Started" Then SvcObj.StartService
    Next SvcIdx
    
    
End Sub

Private Sub BtnAdd_Click()

    FrmServices.Show
    
End Sub


Private Sub BtnAddSchedule_Click()

    Dim FlgFound As Boolean
    
    If IsDate(CmbHour.Text & ":" & CmbMin.Text) Then
        For X = 0 To LstSchedule.ListCount - 1
            If StrComp(LstSchedule.List(X), CmbHour.Text & ":" & CmbMin.Text, vbTextCompare) = 0 Then
                FlgFound = True
                Exit For
            End If
        Next X
        If Not FlgFound Then LstSchedule.AddItem CmbHour.Text & ":" & CmbMin.Text
    End If
        
End Sub

Private Sub BtnOK_Click()
    
    If Dir(App.Path & "\Services.dat") <> "" Then Kill (App.Path & "\Services.dat")
    Open App.Path & "\Services.dat" For Random As #1 Len = Len(RecHolder)
    
    For X = 0 To LstServices.ListCount - 1
        With RecHolder
            .ServiceName = ArrServices(0, X)
            .ServiceDispName = ArrServices(1, X)
        End With
        Put #1, , RecHolder
    Next X
    Close #1
    
    If Dir(App.Path & "\Schedules.dat") <> "" Then Kill (App.Path & "\Schedules.dat")
    Open App.Path & "\Schedules.dat" For Random As #2 Len = Len(SchHolder)
    
    For X = 0 To LstSchedule.ListCount - 1
        SchHolder.Time = LstSchedule.List(X)
        Put #2, , SchHolder
    Next X
    Close #2
    
SaveSetting "Server Guard", "Settings", "Restart Hour", CmbResHr.ListIndex
SaveSetting "Server Guard", "Settings", "Restart Min", CmbResMin.ListIndex
SaveSetting "Server Guard", "Settings", "Restart Enabled", ChkRestart.Value

    Me.Hide
    
End Sub


Private Sub BtnRemove_Click()

    On Error Resume Next
    For X = LstServices.ListIndex To LstServices.ListCount - 2
        ArrServices(0, X) = ArrServices(0, X + 1)
        ArrServices(1, X) = ArrServices(1, X + 1)
    Next X
    ReDim Preserve ArrServices(1, LstServices.ListCount - 2)
    LstServices.RemoveItem LstServices.ListIndex
    
End Sub

Private Sub BtnRemoveSchedule_Click()

    For X = 0 To LstSchedule.ListCount - 1
        If LstSchedule.Selected(X) Then
            LstSchedule.RemoveItem X
            Exit For
        End If
    Next X
    
End Sub

Private Sub ChkAutoStrt_Click()

    If ChkAutoStrt.Value = 1 Then
        AutoStartAdd
    Else
        AutoStartDelete
    End If
    
End Sub

Private Sub Form_Load()

    Dim Swp     As String
    
    Set SvcObj = New ClsService
    
    TrayIconCreate Me, "Server Guard"

    For X = 0 To 23
        CmbHour.AddItem Format(X, "0#")
        CmbMin.AddItem Format(X, "0#")
        CmbResHr.AddItem Format(X, "0#")
        CmbResMin.AddItem Format(X, "0#")
    Next X
    For X = 24 To 59
        CmbMin.AddItem Format(X, "0#")
        CmbResMin.AddItem Format(X, "0#")
    Next X
    CmbHour.ListIndex = 0
    CmbMin.ListIndex = 0
    
    Open App.Path & "\Services.dat" For Random As #1 Len = Len(RecHolder)
    ReDim ArrServices(1, LOF(1) \ Len(RecHolder))
    
    For X = 1 To LOF(1) \ Len(RecHolder)
        Get #1, X, RecHolder
        ArrServices(0, X - 1) = Trim(RecHolder.ServiceName)
        ArrServices(1, X - 1) = Trim(RecHolder.ServiceDispName)
        LstServices.AddItem Trim(RecHolder.ServiceDispName)
    Next X
    Close #1
    
CmbResHr.ListIndex = GetSetting("Server Guard", "Settings", "Restart Hour", 0)
CmbResMin.ListIndex = GetSetting("Server Guard", "Settings", "Restart Min", 0)
ChkRestart.Value = GetSetting("Server Guard", "Settings", "Restart Enabled", 0)

    Open App.Path & "\Schedules.dat" For Random As #2 Len = Len(SchHolder)
    
    For X = 1 To (LOF(2) \ Len(SchHolder))
        Get #2, , SchHolder
        If Trim(SchHolder.Time) <> "" Then LstSchedule.AddItem Trim(SchHolder.Time)
    Next X
    Close #2
    ChkAutoStrt = CInt(Abs(IsAutoStart))
    CheckServices
    
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Msg = X / Screen.TwipsPerPixelX
    If Msg = WM_LBUTTONDBLCLK Then
        Me.Show
    ElseIf Msg = WM_RBUTTONUP Then
        Me.PopupMenu MnuMain, , , , MnuSetup
    End If

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Cancel = 1
    Me.Hide
    
End Sub


Private Sub MnuAbout_Click()

    FrmAbout.Show
    
End Sub

Private Sub MnuExit_Click()

    Msg = "Programdan çýkmak istediðinizden emin misiniz?"
    Msg = Msg & vbCrLf & vbCrLf
    Msg = Msg & "Program sonlandýrýldýðýnda çizelgeler ve hizmet kontrolleri"
    Msg = Msg & vbCrLf & "devre dýþý kalacak..."
    Msg = Msg & vbCrLf & "Devam etmek istiyor musunuz?"
    If MsgBox(Msg, vbQuestion + vbYesNo + vbDefaultButton2, "O N A Y") = vbYes Then
        Shell_NotifyIcon NIM_DELETE, TrayIcon
        End
    End If
    
End Sub

Private Sub MnuServices_Click()

    With FrmServices
        .Label1.Visible = True
        .BtnOK.Enabled = False
        .Show
    End With
    
End Sub

Private Sub MnuSetup_Click()

    Me.Show
    
End Sub

Private Sub Timer1_Timer()

    'Static Mins As Integer
    
    For X = 0 To LstSchedule.ListCount - 1
        If StrComp(Format(Now, "hh:mm"), LstSchedule.List(X)) = 0 Then
            CheckServices
            'MsgBox Now
            'InitiateShutdownMachine GetMyMachineName, True, True, True, 30, "Çizelge uyarýnca sýfýrlama"
        End If
    Next X
    
    If StrComp(Format(Now, "hh:mm"), CmbResHr.Text & ":" & CmbResMin.Text) = 0 Then
        InitiateShutdownMachine GetMyMachineName, True, True, True, 30, "Çizelge uyarýnca sýfýrlama"
    End If
    
    'Mins = Mins + 1
    'If Mins = 30 Then
    '    Mins = 0
    '    CheckServices
    'End If
    
End Sub


