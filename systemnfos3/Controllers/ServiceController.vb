Public Class ServiceController

    Private Property WMI As WMIController

    ''' <summary>
    ''' Connects to the local computer for service control
    ''' </summary>
    ''' <remarks></remarks>
    Sub New(Optional WMI As WMIController = Nothing)
        If WMI Is Nothing Then
            WMI = New WMIController(".")
        End If

        Me.WMI = WMI
    End Sub

    Public Function [Stop](serviceName As String) As ServiceError
        For Each service As ManagementObject In WMI.Query(String.Format("SELECT * FROM Win32_Service WHERE Name=""{0}""", serviceName))
            Return service.InvokeMethod("StopService", Nothing)
        Next

        Return ServiceError.Success
    End Function

    Public Function Start(serviceName As String) As ServiceError
        For Each service As ManagementObject In WMI.Query(String.Format("SELECT * FROM Win32_Service WHERE Name=""{0}""", serviceName))
            Return service.InvokeMethod("StartService", Nothing)
        Next

        Return ServiceError.Success
    End Function

    Public Sub ChangeStartupType(serviceName As String, startupType As ServiceStartupType)
        For Each service As ManagementObject In WMI.Query(String.Format("SELECT * FROM Win32_Service WHERE Name=""{0}""", serviceName))
            Dim startMode As ManagementBaseObject = service.GetMethodParameters("ChangeStartMode")

            Select Case startupType
                Case ServiceStartupType.Auto
                    startMode.SetPropertyValue("StartMode", "Automatic")
                Case ServiceStartupType.Manual
                    startMode.SetPropertyValue("StartMode", "Manual")
                Case ServiceStartupType.Disabled
                    startMode.SetPropertyValue("StartMode", "Disabled")
            End Select

            service.InvokeMethod("ChangeStartMode", startMode, Nothing)
        Next
    End Sub

    Public Function CheckDependentServices(serviceName As String) As String()
        Dim results As New List(Of String)
        Dim services As ManagementObjectCollection = WMI.Query(String.Format("ASSOCIATORS OF {{Win32_Service.Name='{0}'}} WHERE AssocClass=Win32_DependentService Role=Antecedent", serviceName))
        If services.Count > 0 Then
            For Each dependentService As ManagementObject In services
                results.Add(dependentService.Properties("Name").Value)
            Next
        End If

        Return results.ToArray()
    End Function

    ''' <summary>
    ''' Wait for a service to reach a specific state within a specified timeframe
    ''' </summary>
    ''' <param name="serviceName">Name of the Service</param>
    ''' <param name="state">Desired Status Name</param>
    ''' <param name="timeOut">How many iterations to wait. For indefinite iterations, use -1. To check if a service is in a specific status, use 0.</param>
    ''' <returns>Boolean determines whether the service has reached the desired status within the amount of iterations specified.</returns>
    ''' <remarks></remarks>
    Public Function WaitForService(serviceName As String, state As ServiceState, Optional timeOut As Integer = -1) As Boolean
        Dim attemptCount As Integer = 0

        Do Until QueryState(serviceName) = state
            attemptCount += 1
            If timeOut <> -1 AndAlso attemptCount > timeOut Then
                Return False
            End If

            Threading.Thread.Sleep(900)
        Loop

        Return True
    End Function

    Public Function QueryState(serviceName As String) As ServiceState
        Dim serviceState As String = WMI.GetPropertyValue(WMI.Query(String.Format("SELECT State FROM Win32_Service WHERE Name=""{0}""", serviceName)), "State").ToUpper().Trim()

        Select Case serviceState
            Case "RUNNING"
                Return ServiceController.ServiceState.Running
            Case "STOPPING"
                Return ServiceController.ServiceState.Stopping
            Case "STARTING"
                Return ServiceController.ServiceState.Starting
            Case "STOPPED"
                Return ServiceController.ServiceState.Stopped
        End Select

        Return ServiceController.ServiceState.Unknown
    End Function

    Public Function QueryStartupType(serviceName As String) As ServiceStartupType
        Select Case WMI.GetPropertyValue(WMI.Query(String.Format("SELECT StartMode FROM Win32_Service Where Name=""{0}""", serviceName)), "StartMode").ToUpper().Trim()
            Case "AUTO"
                Return ServiceStartupType.Auto
            Case "MANUAL"
                Return ServiceStartupType.Manual
            Case "DISABLED"
                Return ServiceStartupType.Disabled
        End Select

        Return ServiceStartupType.Unknown
    End Function

    Public Enum ServiceState
        Running = 1
        Stopping = 2
        Starting = 3
        Stopped = 4
        Unknown = 5
    End Enum

    Public Enum ServiceStartupType
        Auto = 0
        Manual = 1
        Disabled = 2
        Unknown = 3
    End Enum

    Public Enum ServiceError
        Success = 0
        NotSupported = 1
        AccessDenied = 2
        DependentServicesRunning = 3
        InvalidServicecontrol = 4
        ServiceCannotAcceptControl = 5
        ServiceNotActive = 6
        ServiceRequestTimeout = 7
        UnknownFailure = 8
        PathNotFound = 9
        ServiceAlreadyStopped = 10
        ServiceDatabaseLocked = 11
        ServiceDependencyDeleted = 12
        ServiceDependencyFailure = 13
        ServiceDisabled = 14
        ServiceLogonFailed = 15
        ServiceMarkedForDeletion = 16
        ServiceNoThread = 17
        StatusCircularDependency = 18
        StatusDuplicateName = 19
        StatusInvalidName = 20
        StatusInvalidParameter = 21
        StatusInvalidServiceAccount = 22
        StatusServiceExists = 23
        ServiceAlreadyPaused = 24
    End Enum

End Class