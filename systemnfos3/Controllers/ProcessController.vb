Imports System.Reflection

Public Class ProcessController

    Private Property WMI As WMIController


    ''' <summary>
    ''' Connects to the local computer for process control
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub New(Optional WMI As WMIController = Nothing)
        If WMI Is Nothing Then
            WMI = New WMIController(".")
        End If

        Me.WMI = WMI
    End Sub

    Public Function StopProcess(processID As String) As ProcessError
        Try
            For Each process As ManagementObject In WMI.Query(String.Format("SELECT * FROM Win32_Process WHERE ProcessID=""{0}""", processID))
                Return process.InvokeMethod("Terminate", Nothing)
            Next

            Return ProcessError.Success
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try

        Return ProcessError.UnknownFailure
    End Function

    Public Function GetProcesses(processName As String) As ManagementObject()
        Try
            Dim processes As New List(Of ManagementObject)
            For Each process As ManagementObject In WMI.Query(String.Format("SELECT * FROM Win32_Process WHERE Name=""{0}""", processName))
                processes.Add(process)
            Next

            Return processes.ToArray()
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' Wait for a process to reach a specific state within a specified timeframe
    ''' </summary>
    ''' <param name="processName">Name of the process</param>
    ''' <param name="condition">Desired status</param>
    ''' <param name="timeOut">How many iterations to wait. For indefinite iterations, use -1. To check if a process is in a specific status, use 0.</param>
    ''' <returns>Boolean determines whether the process has reached the desired status within the amount of iterations specified.</returns>
    ''' <remarks></remarks>
    Public Function WaitForProcess(processName As String, condition As ProcessCondition, Optional timeOut As Integer = -1) As Boolean
        Dim attemptCount As Integer = 0
        Do Until QueryProcessState(processName) = condition
            attemptCount += 1

            If timeOut <> -1 Then
                If attemptCount > timeOut Then Return False
            End If

            Threading.Thread.Sleep(900)
        Loop

        Return True
    End Function

    Public Function QueryProcessState(processName As String) As ProcessCondition
        Try
            Select Case WMI.Query(String.Format("SELECT * FROM Win32_Process WHERE Name=""{0}""", processName)).Count
                Case = 1
                    Return ProcessCondition.SingleRunning
                Case > 1
                    Return ProcessCondition.MultipleRunning
                Case = 0
                    Return ProcessCondition.Stopped
                Case Else
                    Return ProcessCondition.Unknown
            End Select

        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try

        Return ProcessCondition.Unknown
    End Function

    Public Enum ProcessError
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
        ProcessAlreadyStopped = 10
        ProcessDatabaseLocked = 11
        ProcessLogonFailed = 15
        StatusDuplicateName = 19
        StatusInvalidName = 20
        StatusInvalidParameter = 21
        StatusInvalidServiceAccount = 22
        StatusProcessExists = 23
        ProcessAlreadyPaused = 24
    End Enum

    Public Enum ProcessCondition
        SingleRunning = 1
        MultipleRunning = 2
        Stopped = 4
        Unknown = 5
    End Enum

End Class