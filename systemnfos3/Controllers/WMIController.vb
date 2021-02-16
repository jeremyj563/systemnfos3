Imports System.ComponentModel
Imports System.Reflection
Imports System.Threading

Public NotInheritable Class WMIController

    Public Property Architecture As ComputerArchitectures = ComputerArchitectures.X86
    Public Property X64Scope As ManagementScope
    Public Property X86Scope As ManagementScope
    Public Property BitLockerScope As ManagementScope
    Public Property RegularScope As ManagementScope
    Public Property TPMScope As ManagementScope
    Public Property LDAPScope As ManagementScope
    Public Property ConnectedTime As Date
    Public ReadOnly Property ComputerName As String

    Public Sub New(computerName As String, Optional scope As ManagementScopes = ManagementScopes.All, Optional async As Boolean = False)
        Me.ComputerName = computerName

        Connect(computerName, scope, async)
    End Sub

    Public Sub Connect(computerName As String, scope As ManagementScopes, async As Boolean)
        Dim options As New ConnectionOptions With
            {
                .Impersonation = ImpersonationLevel.Impersonate,
                .Authentication = AuthenticationLevel.PacketPrivacy,
                .EnablePrivileges = True,
                .Timeout = New TimeSpan(0, 0, 0, 2, 0)
            }

        Select Case scope
            Case ManagementScopes.Regular
                SetManagementScopeRegular(computerName, options, async)

            Case ManagementScopes.LDAP
                SetManagementScopeLDAP(computerName, options)

            Case ManagementScopes.X86
                SetManagementScopeX86(computerName, options)

            Case ManagementScopes.X64
                SetManagementScopeX64(computerName, options)

            Case ManagementScopes.BitLocker
                SetManagementScopeBitLocker(computerName, options)

            Case ManagementScopes.TPM
                SetManagementScopeTPM(computerName, options)

            Case ManagementScopes.All
                SetManagementScopeRegular(computerName, options, async)
                SetManagementScopeLDAP(computerName, options)
                SetManagementScopeX86(computerName, options)
                SetManagementScopeX64(computerName, options)
                SetManagementScopeBitLocker(computerName, options)
                SetManagementScopeTPM(computerName, options)

        End Select

        Me.ConnectedTime = Now
    End Sub

    Public Function GetPropertyValue(managementObjects As ManagementObjectCollection, [property] As String)
        Dim propertyValue As Object = Nothing

        If Not String.IsNullOrWhiteSpace([property]) AndAlso managementObjects IsNot Nothing Then
            Try
                Dim managementObject As ManagementObject = managementObjects.Cast(Of ManagementObject)().FirstOrDefault()
                If managementObject IsNot Nothing Then
                    propertyValue = managementObject.GetPropertyValue([property])
                End If
            Catch ex As Exception
                ' Return null
            End Try
        End If

        Return propertyValue
    End Function

    ''' <summary>
    ''' Used to return a collection of WMI properties.
    ''' </summary>
    ''' <param name="queryText">WQL Query for the search</param>
    ''' <param name="scope">Provide specialized search scope for the query</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function Query(queryText As String, Optional scope As ManagementScope = Nothing) As ManagementObjectCollection
        If scope Is Nothing Then scope = Me.RegularScope

        Dim result As ManagementObjectCollection = Nothing
        Try
            Dim searcher As New ManagementObjectSearcher(scope.Path.ToString(), queryText)
            result = searcher.Get()
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return result
    End Function

    Private Sub SetManagementScopeTPM(computerName As String, options As ConnectionOptions)
        Me.TPMScope = New ManagementScope()

        If GetOperatingSystemType() = OperatingSystemTypes.Client Then
            Try
                With Me.TPMScope
                    .Path.Server = computerName
                    .Path.NamespacePath = "root\cimv2\Security\MicrosoftTpm"
                    .Path.ClassName = "Win32_Tpm"
                    .Options = options
                    .Connect()
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub

    Private Sub SetManagementScopeBitLocker(computerName As String, options As ConnectionOptions)
        Me.BitLockerScope = New ManagementScope()

        If GetOperatingSystemType() = OperatingSystemTypes.Client Then
            Try
                With Me.BitLockerScope
                    .Path.Server = computerName
                    .Path.NamespacePath = "root\cimv2\Security\MicrosoftVolumeEncryption"
                    .Path.ClassName = "Win32_EncryptableVolume"
                    .Options = options
                    .Connect()
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub

    Private Sub SetManagementScopeX64(computerName As String, options As ConnectionOptions)
        Me.X64Scope = New ManagementScope()

        If GetComputerArchitecture() = ComputerArchitectures.X64 Then
            Me.Architecture = ComputerArchitectures.X64

            Try
                With Me.X64Scope
                    .Path.Server = computerName
                    .Path.NamespacePath = "\root\default"
                    .Options = options
                    .Options.Context.Add("__ProviderArchitecture", ComputerArchitectures.X64)
                    .Connect()
                End With
            Catch ex As Exception
                Throw ex
            End Try
        End If
    End Sub

    Private Sub SetManagementScopeX86(computerName As String, options As ConnectionOptions)
        Me.X86Scope = New ManagementScope()

        Try
            With Me.X86Scope
                .Path.Server = computerName
                .Path.NamespacePath = "\root\default"
                .Options = options
                .Options.Context.Add("__ProviderArchitecture", ComputerArchitectures.X86)
                .Connect()
            End With
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub SetManagementScopeRegular(computerName As String, options As ConnectionOptions, async As Boolean, Optional asyncTimeout As TimeSpan = Nothing)
        Me.RegularScope = New ManagementScope()

        Try
            With Me.RegularScope
                .Path.Server = computerName
                .Path.NamespacePath = "\root\cimv2"
                .Options = options
            End With

            If async Then
                ' Default timeout is 500 milliseconds
                If asyncTimeout = Nothing Then
                    asyncTimeout = New TimeSpan(0, 0, 0, 0, 500)
                End If

                ' Initiate connection on a background thread
                Dim connectionWorker As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
                AddHandler connectionWorker.DoWork, Sub()
                                                        Try
                                                            Me.RegularScope.Connect()
                                                        Catch ex As Exception
                                                            ' Silently fail
                                                            Exit Sub
                                                        End Try
                                                    End Sub

                connectionWorker.RunWorkerAsync()
                Thread.Sleep(asyncTimeout)

                If Not Me.RegularScope.IsConnected Then
                    ' Connection could not be esablished within timeout so abort
                    connectionWorker.CancelAsync()
                    Throw New Exception($"Failed to connect to RPC server on {computerName}")
                End If
            Else
                Me.RegularScope.Connect()
            End If

        Catch ex As Exception
            Throw ex
        End Try

        If GetComputerArchitecture() = ComputerArchitectures.X64 Then
            Me.Architecture = ComputerArchitectures.X64
        End If
    End Sub

    Private Sub SetManagementScopeLDAP(computerName As String, options As ConnectionOptions)
        Me.LDAPScope = New ManagementScope()

        Try
            With Me.LDAPScope
                .Path.Server = computerName
                .Path.NamespacePath = "\root\directory\LDAP"
                .Options = options
                .Connect()
            End With
        Catch ex As Exception
            Throw ex
        End Try

        If Me.RegularScope Is Nothing Then
            SetManagementScopeRegular(computerName, options, async:=False)
        End If

        If GetComputerArchitecture() = ComputerArchitectures.X64 Then
            Me.Architecture = ComputerArchitectures.X64
        End If
    End Sub

    Private Function GetComputerArchitecture() As ComputerArchitectures
        Dim architecture As ComputerArchitectures = Nothing

        Dim architectureText As String = GetPropertyValue(Query("SELECT AddressWidth FROM Win32_Processor"), "AddressWidth")
        Select Case architectureText
            Case ComputerArchitectures.X64.ToString("D")
                architecture = ComputerArchitectures.X64
            Case ComputerArchitectures.X86.ToString("D")
                architecture = ComputerArchitectures.X86
        End Select

        Return architecture
    End Function

    Private Function GetOperatingSystemType() As OperatingSystemTypes
        Dim osType As OperatingSystemTypes = Nothing

        Dim osTypeText As String = GetPropertyValue(Query("SELECT ProductType FROM Win32_OperatingSystem"), "ProductType")
        Select Case osTypeText
            Case OperatingSystemTypes.Client.ToString("D")
                osType = OperatingSystemTypes.Client
            Case OperatingSystemTypes.DC.ToString("D")
                osType = OperatingSystemTypes.DC
            Case OperatingSystemTypes.Host.ToString("D")
                osType = OperatingSystemTypes.Host
        End Select

        Return osType
    End Function

    Public Enum ComputerArchitectures
        X86 = 32
        X64 = 64
    End Enum

    Public Enum ManagementScopes
        All
        Regular
        X86
        X64
        BitLocker
        TPM
        LDAP
    End Enum

    Private Enum OperatingSystemTypes
        Client = 1
        DC = 2
        Host = 3
    End Enum

End Class
