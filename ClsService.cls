VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsService"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim svcName As String

Enum ServiceStartType
    svcBootStart = &H0
    svcSystemStart = &H1
    svcAutoStart = &H2
    svcDemandStart = &H3
    svcDisabled = &H4
End Enum

Public Function StartService() As Boolean

    StartService = StartSvc(svcName)

End Function

Public Function StopService() As Boolean

    StopService = StopSvc(svcName)

End Function

Public Function PauseService() As Boolean

    PauseService = PauseSvc(svcName)

End Function

Public Function ResumeService() As Boolean

    ResumeService = ResumeSvc(svcName)

End Function

Public Function QueryService() As String

    QueryService = EnumSvcState(QuerySvc(svcName))

End Function

Public Property Get ServiceName() As String

    ServiceName = svcName

End Property

Public Property Let ServiceName(ByVal vSvcName As String)

    svcName = vSvcName

End Property

Public Function SetStartType(serviceStart As ServiceStartType) As Boolean

    SetStartType = SetSvcStartType(svcName, serviceStart)

End Function

