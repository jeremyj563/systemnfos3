Imports System.ComponentModel
Imports System.IO
Imports System.Reflection
Imports System.Threading
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Public Class RemoteTools

    Public ReadOnly Property IsBusy As Boolean
        Get
            Return Me.BackgroundThread.IsBusy
        End Get
    End Property

    Private Property NumberOfComputers As Integer = 1
    Private Property BackgroundThread As New BackgroundWorker()
    Private Property RemoteToolBackgroundThread As Thread
    Private Property ComputerContext As ComputerControl = Nothing
    Private Property TabPage As TabPage = Nothing
    Private Property WMIObjects As ManagementObject() = Nothing
    Private Property AdminShare As String = Nothing
    Private Property CustomAction As String = Nothing
    Private Property ShutdownTimeout As Integer = 0

    Public Event WorkCompleted As EventHandler

#Region " Constructors "

    Public Sub New(remoteTool As RemoteTools, computerContext As ComputerControl)
        ' The old constructor
        Me.ComputerContext = computerContext

        Select Case remoteTool
            Case RemoteTools.RemoteAssistance
                AddHandler BackgroundThread.DoWork, AddressOf RemoteAssistance

            Case RemoteTools.RemoteDesktop
                AddHandler BackgroundThread.DoWork, AddressOf RemoteDesktop

            Case RemoteTools.RemoteControlViewer
                AddHandler BackgroundThread.DoWork, AddressOf RemoteControlViewer

            Case RemoteTools.RemoteRegistry
                Me.RemoteToolBackgroundThread = New Thread(AddressOf RemoteRegistry)
                Me.RemoteToolBackgroundThread.SetApartmentState(ApartmentState.STA)

            Case RemoteTools.AdminShare
                AddHandler BackgroundThread.DoWork, AddressOf InitializeAdminShare

            Case RemoteTools.PsExec
                Me.RemoteToolBackgroundThread = New Thread(AddressOf PsExecLaunchInteractiveCommandPrompt)
                Me.RemoteToolBackgroundThread.SetApartmentState(ApartmentState.STA)

            Case RemoteTools.ComputerManagement
                AddHandler BackgroundThread.DoWork, AddressOf ComputerManagement

            Case RemoteTools.GroupPolicyEditor
                AddHandler BackgroundThread.DoWork, AddressOf RemoteGroupPolicyEditor

            Case RemoteTools.SetDescription
                AddHandler BackgroundThread.DoWork, AddressOf SetDescription

            Case RemoteTools.SetLocation
                AddHandler BackgroundThread.DoWork, AddressOf SetLocation

            Case RemoteTools.ToggleBitLocker
                AddHandler BackgroundThread.DoWork, AddressOf ToggleBitLocker

            Case RemoteTools.Restart
                AddHandler BackgroundThread.DoWork, AddressOf RemoteRestart

            Case RemoteTools.NoPromptRestart
                AddHandler BackgroundThread.DoWork, AddressOf NoPromptRestart

            Case RemoteTools.Logoff
                AddHandler BackgroundThread.DoWork, AddressOf RemoteLogoff

            Case RemoteTools.PrinterAdd
                AddHandler BackgroundThread.DoWork, AddressOf PrinterAdd

            Case RemoteTools.PrinterAddCab
                Me.RemoteToolBackgroundThread = New Thread(AddressOf PrinterAddCab)
                Me.RemoteToolBackgroundThread.SetApartmentState(ApartmentState.STA)

            Case RemoteTools.PrinterNewCab
                Me.RemoteToolBackgroundThread = New Thread(AddressOf PrinterNewCab)
                Me.RemoteToolBackgroundThread.SetApartmentState(ApartmentState.STA)

            Case RemoteTools.PrinterRename
                AddHandler BackgroundThread.DoWork, AddressOf PrinterRename

            Case RemoteTools.PrinterSetDefault
                AddHandler BackgroundThread.DoWork, AddressOf PrinterSetDefault

            Case RemoteTools.PrinterOpenQueue
                AddHandler BackgroundThread.DoWork, AddressOf PrinterOpenQueue

            Case RemoteTools.PrinterOpenProperties
                AddHandler BackgroundThread.DoWork, AddressOf PrinterOpenProperties

            Case RemoteTools.PrinterEditPort
                AddHandler BackgroundThread.DoWork, AddressOf PrinterEditPort

            Case RemoteTools.PrinterDeleteAnyType
                AddHandler BackgroundThread.DoWork, AddressOf PrinterDeleteAnyType

            Case RemoteTools.ServiceStart
                AddHandler BackgroundThread.DoWork, AddressOf ServiceStart

            Case RemoteTools.ServiceStop
                AddHandler BackgroundThread.DoWork, AddressOf ServiceStop

            Case RemoteTools.ServiceSetAuto
                AddHandler BackgroundThread.DoWork, AddressOf ServiceSetAuto

            Case RemoteTools.ServiceSetManual
                AddHandler BackgroundThread.DoWork, AddressOf ServiceSetManual

            Case RemoteTools.ServiceSetDisable
                AddHandler BackgroundThread.DoWork, AddressOf ServiceSetDisabled

            Case RemoteTools.ServiceRestart
                AddHandler BackgroundThread.DoWork, AddressOf ServiceRestart

            Case RemoteTools.CustomAction
                Me.RemoteToolBackgroundThread = New Thread(AddressOf RunCustomAction)
                Me.RemoteToolBackgroundThread.SetApartmentState(ApartmentState.STA)

            Case RemoteTools.EnableComputers
                AddHandler BackgroundThread.DoWork, AddressOf EnableComputers

            Case RemoteTools.DisableComputers
                AddHandler BackgroundThread.DoWork, AddressOf DisableComputers

            Case RemoteTools.TimedRestart
                AddHandler BackgroundThread.DoWork, AddressOf TimedRemoteRestart

            Case RemoteTools.ProcessStop
                AddHandler BackgroundThread.DoWork, AddressOf ProcessStop

        End Select
    End Sub

    Public Sub New(adminShareAndCustomActionOnly As RemoteTools, sender As Object, driveLetterToOpenOrCustomActionToLoad As String)
        ' The constructor for opening a remote share
        Me.New(adminShareAndCustomActionOnly, sender)
        Select Case adminShareAndCustomActionOnly
            Case RemoteTools.AdminShare
                AdminShare = driveLetterToOpenOrCustomActionToLoad
            Case RemoteTools.CustomAction
                CustomAction = driveLetterToOpenOrCustomActionToLoad
            Case Else
                Exit Sub
        End Select
    End Sub

    Public Sub New(printerOrServiceTool As RemoteTools, sender As Object, managementObject As ManagementObject())
        ' The constructor for using a printer tool or service tool
        Me.New(printerOrServiceTool, sender)
        Me.WMIObjects = managementObject
    End Sub

    Public Sub BeginWork()
        If RemoteToolBackgroundThread Is Nothing Then
            Me.BackgroundThread.RunWorkerAsync()
        Else
            Me.RemoteToolBackgroundThread.Start()
        End If
    End Sub

#End Region

#Region " Tools "

    Public Enum RemoteTools
        RemoteAssistance = 1
        RemoteDesktop = 2
        RemoteRegistry = 3
        AdminShare = 4
        PsExec = 5
        ComputerManagement = 6
        GroupPolicyEditor = 7
        ResultantSetOfPolicy = 8
        CACExempt = 9
        SetDescription = 10
        SetLocation = 11
        ToggleBitLocker = 12
        Restart = 13
        NoPromptRestart = 14
        Logoff = 15
        PrinterAdd = 16
        PrinterAddCab = 17
        PrinterNewCab = 18
        PrinterRename = 19
        PrinterSetDefault = 20
        PrinterOpenQueue = 21
        PrinterOpenProperties = 22
        PrinterEditPort = 23
        PrinterDeleteAnyType = 24
        ServiceStart = 25
        ServiceStop = 26
        ServiceSetAuto = 27
        ServiceSetManual = 28
        ServiceSetDisable = 29
        ServiceRestart = 30
        TEMP_WriteToScanList = 31
        SubmitHBSSScan = 32
        CustomAction = 33
        EnableComputers = 34
        DisableComputers = 35
        TimedRestart = 36
        ProcessStop = 37
        RemoteControlViewer = 38
    End Enum

    Private Structure RegistryPaths
        Public Shared ReadOnly Property ScForceOption As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Public Shared ReadOnly Property ConsentPromptBehaviorAdmin As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Public Shared ReadOnly Property ConsentPromptBehaviorUser As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Public Shared ReadOnly Property EnableSecureCredentialPrompting As String = "Software\Microsoft\Windows\CurrentVersion\Policies\CredUI"
        Public Shared ReadOnly Property EnableUIADesktopToggle As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
        Public Shared ReadOnly Property PromptOnSecureDesktop As String = "Software\Microsoft\Windows\CurrentVersion\Policies\System"
    End Structure

#Region " Main "

#Region "Remote Assistance"

    Private Structure RemoteAssistanceRegistryValues
        Public Property ScForceOption As Integer
        Public Property ConsentPromptBehaviorAdmin As Integer
        Public Property ConsentPromptBehaviorUser As Integer
        Public Property EnableUIADesktopToggle As Integer
        Public Property EnableSecureCredentialPrompting As Integer
        Public Property PromptOnSecureDesktop As Integer
    End Structure

    Private Function RemoteAssistanceGetRegistryValues(registry As RegistryController) As RemoteAssistanceRegistryValues
        Return New RemoteAssistanceRegistryValues() With
        {
            .ScForceOption = registry.GetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption)),
            .ConsentPromptBehaviorAdmin = registry.GetKeyValue(RegistryPaths.ConsentPromptBehaviorAdmin, NameOf(RegistryPaths.ConsentPromptBehaviorAdmin)),
            .ConsentPromptBehaviorUser = registry.GetKeyValue(RegistryPaths.ConsentPromptBehaviorUser, NameOf(RegistryPaths.ConsentPromptBehaviorUser)),
            .EnableSecureCredentialPrompting = registry.GetKeyValue(RegistryPaths.EnableSecureCredentialPrompting, NameOf(RegistryPaths.EnableSecureCredentialPrompting)),
            .EnableUIADesktopToggle = registry.GetKeyValue(RegistryPaths.EnableUIADesktopToggle, NameOf(RegistryPaths.EnableUIADesktopToggle)),
            .PromptOnSecureDesktop = registry.GetKeyValue(RegistryPaths.PromptOnSecureDesktop, NameOf(RegistryPaths.PromptOnSecureDesktop))
        }
    End Function

    Private Sub RemoteAssistanceSetRegistryValues(registry As RegistryController, values As RemoteAssistanceRegistryValues)
        registry.SetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption), values.ScForceOption)
        registry.SetKeyValue(RegistryPaths.ConsentPromptBehaviorAdmin, NameOf(RegistryPaths.ConsentPromptBehaviorAdmin), values.ConsentPromptBehaviorAdmin)
        registry.SetKeyValue(RegistryPaths.ConsentPromptBehaviorUser, NameOf(RegistryPaths.ConsentPromptBehaviorUser), values.ConsentPromptBehaviorUser)
        registry.SetKeyValue(RegistryPaths.EnableUIADesktopToggle, NameOf(RegistryPaths.EnableUIADesktopToggle), values.EnableUIADesktopToggle)
        registry.SetKeyValue(RegistryPaths.EnableSecureCredentialPrompting, NameOf(RegistryPaths.EnableSecureCredentialPrompting), values.EnableSecureCredentialPrompting)
        registry.SetKeyValue(RegistryPaths.PromptOnSecureDesktop, NameOf(RegistryPaths.PromptOnSecureDesktop), values.PromptOnSecureDesktop)
    End Sub

    Private Sub RemoteAssistance()
        Try
            TryWriteMessage(String.Format("Initializing Remote Assistance on {0}", ComputerContext.Computer.ConnectionString), Color.Blue)

            ' Save original registry values so they can be restored when the tool terminates
            Dim x86Registry As New RegistryController(ComputerContext.WMI.X86Scope)
            Dim x86OriginalRegistryValues As RemoteAssistanceRegistryValues = RemoteAssistanceGetRegistryValues(x86Registry)

            Dim x64Registry As New RegistryController(Me.ComputerContext.WMI.X64Scope)
            Dim x64OriginalRegistryValues As RemoteAssistanceRegistryValues = RemoteAssistanceGetRegistryValues(x64Registry)

            ' Set temporary registry values for this session only
            TryWriteMessage("Changing registry values", Color.Blue)

            Dim temporaryRegistryValues As New RemoteAssistanceRegistryValues With
            {
                .ScForceOption = 0,
                .ConsentPromptBehaviorAdmin = 1,
                .ConsentPromptBehaviorUser = 1,
                .EnableUIADesktopToggle = 1,
                .EnableSecureCredentialPrompting = 0,
                .PromptOnSecureDesktop = 0
            }

            RemoteAssistanceSetRegistryValues(x86Registry, temporaryRegistryValues)
            If Me.ComputerContext.WMI.Architecture = WMIController.ComputerArchitectures.X64 Then
                RemoteAssistanceSetRegistryValues(x64Registry, temporaryRegistryValues)
            End If


            ' Start the remote tool session
            TryWriteMessage("Starting Remote Assistance...", Color.Blue)
            Process.Start("msra.exe", String.Format("/OFFERRA {0}", Me.ComputerContext.Computer.ConnectionString)).WaitForExit()


            ' The session has ended so restore original registry values
            TryWriteMessage("Remote assistance has been terminated... Restoring registry values to initial settings", Color.Blue)

            RemoteAssistanceSetRegistryValues(x86Registry, x86OriginalRegistryValues)
            If Me.ComputerContext.WMI.Architecture = WMIController.ComputerArchitectures.X64 Then
                RemoteAssistanceSetRegistryValues(x64Registry, x64OriginalRegistryValues)
            End If

            TryWriteMessage(String.Format("Terminated Remote Assistance on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

#End Region

    Private Sub RemoteDesktop()
        Try
            TryWriteMessage(String.Format("Initializing Remote Desktop on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            ' Save original registry values so they can be restored when the tool terminates
            Dim x86Registry As New RegistryController(Me.ComputerContext.WMI.X86Scope)
            Dim originalScForceOption_x86 As Integer = x86Registry.GetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption))

            Dim x64Registry As New RegistryController(Me.ComputerContext.WMI.X64Scope)
            Dim originalScForceOption_x64 As Integer = x64Registry.GetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption))

            TryWriteMessage("Changing registry values", Color.Blue)

            ' Set temporary registry values for this session only
            x86Registry.SetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption), 0)
            If Me.ComputerContext.WMI.Architecture = WMIController.ComputerArchitectures.X64 Then
                x64Registry.SetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption), 0)
            End If

            ' Ensure any needed services are running
            TryWriteMessage("Checking the availability of the Terminal Service", Color.Blue)

            Dim services As New ServiceController(Me.ComputerContext.WMI)
            If services.QueryState("TermService") <> ServiceController.ServiceState.Running Then
                services.Start("TermService")
                If Not services.WaitForService("TermService", ServiceController.ServiceState.Running, 10) Then
                    TryWriteMessage("Unable to start Remote Desktop. The Terminal Service on the Remote computer will not start in a timely manner.", Color.Red)
                    Exit Sub
                End If
            End If

            ' Start the remote tool session
            TryWriteMessage("Starting Remote Desktop", Color.Blue)
            Process.Start("mstsc.exe", String.Format("/v:{0}", Me.ComputerContext.Computer.ConnectionString)).WaitForExit()
            Thread.Sleep(1000)

            ' Wait until there are no running instances of Remote Desktop
            Do Until Process.GetProcessesByName("mstsc").Count = 0
                Thread.Sleep(1000)
            Loop

            ' The session has ended so restore original registry values
            TryWriteMessage("Remote Desktop has been terminated... Restoring registry values to initial settings", Color.Blue)

            x86Registry.SetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption), originalScForceOption_x86)
            If Me.ComputerContext.WMI.Architecture = WMIController.ComputerArchitectures.X64 Then
                x64Registry.SetKeyValue(RegistryPaths.ScForceOption, NameOf(RegistryPaths.ScForceOption), originalScForceOption_x64)
            End If

            TryWriteMessage(String.Format("Terminated Remote Desktop on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage(String.Format("Unable to start Remote Desktop {0}", ex.Message), Color.Red)
        End Try
    End Sub

    Private Sub RemoteControlViewer()
        Try
            Dim PathToRemoteControlViewer As String = Nothing
            If ComputerContext.WMI.Architecture = WMIController.ComputerArchitectures.X64 Then
                PathToRemoteControlViewer = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86), "Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe")
            ElseIf ComputerContext.WMI.Architecture = WMIController.ComputerArchitectures.X86 Then
                PathToRemoteControlViewer = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "Microsoft Configuration Manager\AdminConsole\bin\i386\CmRcViewer.exe")
            End If

            If Not File.Exists(PathToRemoteControlViewer) Then
                TryWriteMessage("Unable to find Remote Control Viewer. Please verify that it has been installed on the local computer.", Color.Red)
                Exit Sub
            End If

            TryWriteMessage(String.Format("Initializing Remote Control Viewer on {0}", ComputerContext.Computer.ConnectionString), Color.Blue)
            Dim ServiceController As New ServiceController(ComputerContext.WMI)

            TryWriteMessage("Checking the availability of the Remote Control Service", Color.Blue)
            If ServiceController.QueryState("CmRcService") <> ServiceController.ServiceState.Running Then
                ServiceController.Start("CmRcService")
                If Not ServiceController.WaitForService("CmRcService", ServiceController.ServiceState.Running, 10) Then
                    TryWriteMessage("Unable to start Remote Control Viewer. The Remote Control Service on the Remote computer will not start in a timely manner.", Color.Red)
                    Exit Sub
                End If
            End If


            ' Launch a new instance of Remote Control Viewer
            TryWriteMessage("Starting Remote Control Viewer", Color.Blue)
            Process.Start(PathToRemoteControlViewer, ComputerContext.Computer.ConnectionString).WaitForExit()
            Thread.Sleep(1000)

            ' Wait until there are no running instances of Remote Control Viewer
            Do Until Process.GetProcessesByName("CmRcViewer").Count = 0
                Thread.Sleep(1000)
            Loop

            TryWriteMessage(String.Format("Terminated Remote Control Viewer on {0}", ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage(String.Format("Unable to start Remote Control Viewer {0}", ex.Message), Color.Red)
        End Try
    End Sub

    Private Sub RemoteRegistry()
        Try
            TryWriteMessage("Initializing Remote Registry", Color.Blue)
            TryWriteMessage("Copying ComputerName to the Clipboard", Color.Blue)

            Clipboard.SetText(Me.ComputerContext.Computer.ConnectionString)

            TryWriteMessage("Starting registry editor", Color.Blue)
            Process.Start("regedit.exe")
            Thread.Sleep(700)


            TryWriteMessage("Sending Keys: Alt+F", Color.Blue)
            SendKeys.SendWait("%f")

            TryWriteMessage("Sending Keys: C", Color.Blue)
            SendKeys.SendWait("c")
            Thread.Sleep(300)

            TryWriteMessage("Sending Keys: Ctrl+V", Color.Blue)
            SendKeys.SendWait("^v")

            TryWriteMessage("Sending Keys: Enter", Color.Blue)
            SendKeys.SendWait("{ENTER}")
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub InitializeAdminShare()
        Try
            TryWriteMessage(String.Format("Initializing {0} Share on {1}", Me.AdminShare, Me.ComputerContext.Computer.ConnectionString), Color.Blue)
            Process.Start("explorer.exe", String.Format("\\{0}\{1}", Me.ComputerContext.Computer.ConnectionString, Me.AdminShare)).WaitForExit()
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PsExecLaunchInteractiveCommandPrompt()
        Try
            TryWriteMessage("Attempting to locate PsExec.exe", Color.Blue)
            [Global].SetPsExecBinaryPath()

            If String.IsNullOrWhiteSpace(My.Settings.PsExecPath) OrElse Not File.Exists(My.Settings.PsExecPath) Then
                TryWriteMessage("Unable to locate a valid copy of PsExec.exe", Color.Red)
                Exit Sub
            Else
                TryWriteMessage(String.Format("Initializing PsExec on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

                Dim psExecProcess As New ProcessStartInfo With
                    {
                        .FileName = "cmd.exe",
                        .Arguments = String.Format(" /c ""{0}"" -s \\{1} cmd.exe", My.Settings.PsExecPath, Me.ComputerContext.Computer.ConnectionString),
                        .Verb = "RunAs",
                        .WorkingDirectory = Path.GetDirectoryName(My.Settings.PsExecPath)
                    }
                Process.Start(psExecProcess).WaitForExit()
            End If

            TryWriteMessage(String.Format("Terminated PsExec on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ComputerManagement()
        Try
            TryWriteMessage(String.Format("Initializing Computer Management on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
            Process.Start("mmc.exe", String.Format("{0} /computer={1}", "compmgmt.msc", Me.ComputerContext.Computer.ConnectionString)).WaitForExit()
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub RemoteGroupPolicyEditor()
        Try
            TryWriteMessage(String.Format("Initializing Remote Group Policy Editor on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
            Process.Start("mmc.exe", String.Format("gpedit.msc /gpcomputer: {0}", Me.ComputerContext.Computer.ConnectionString)).WaitForExit()
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub SetDescription()
        Try
            TryWriteMessage(String.Format("Initializing Set Description on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            Dim currentDescription As String = Me.ComputerContext.Computer.ActiveDirectoryContainer.GetAttribute("Description")
            Dim newDescription As String = InputBox("Enter a new description:", "Set Description", currentDescription)

            If newDescription IsNot currentDescription AndAlso Not String.IsNullOrWhiteSpace(newDescription) Then
                ' Set the ldap 'description' attribute value on the 'computer' class instance
                Me.ComputerContext.Computer.ActiveDirectoryContainer.Description = newDescription

                ' Set the WMI 'Description' property value on the 'Win32_OperatingSystem' class instance
                If Me.ComputerContext.ConnectionStatus = ComputerControl.ConnectionStatuses.Online Then
                    For Each win32_OperatingSystem As ManagementObject In Me.ComputerContext.WMI.Query("SELECT * FROM Win32_OperatingSystem WHERE Primary=TRUE")
                        win32_OperatingSystem("Description") = newDescription
                        win32_OperatingSystem.Put()
                    Next
                End If

                TryWriteMessage(String.Format("Description successfully changed from {0} to {1} on {2}", currentDescription, newDescription, Me.ComputerContext.Computer.ConnectionString), Color.Blue)
                TryWriteMessage("The description will update in the system tool once the change is reflected in LDAP", Color.Blue)
            End If

            TryWriteMessage(String.Format("Terminated Set Description on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub SetLocation()
        Try
            TryWriteMessage(String.Format("Initializing Set Location on {0}", ComputerContext.Computer.ConnectionString), Color.Blue)

            Dim currentLocation As String = Me.ComputerContext.Computer.ActiveDirectoryContainer.PhysicalDeliveryOfficeName
            Dim newLocation As String = InputBox("Enter a new location:", "Set Location", currentLocation)

            If newLocation IsNot currentLocation AndAlso Not String.IsNullOrWhiteSpace(newLocation) Then
                Me.ComputerContext.Computer.ActiveDirectoryContainer.PhysicalDeliveryOfficeName = newLocation

                TryWriteMessage(String.Format("Location successfully changed from {0} to {1} on {2}", currentLocation, newLocation, ComputerContext.Computer.ConnectionString), Color.Blue)

                RaiseEvent WorkCompleted(Me, EventArgs.Empty)
            End If

            TryWriteMessage(String.Format("Terminated Set Location on {0}", ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ToggleBitLocker()
        Try
            TryWriteMessage(String.Format("Initializing Toggle BitLocker on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            If Me.ComputerContext.WMI.BitLockerScope.IsConnected Then
                TryWriteMessage("BitLocker is installed on this computer. Checking status", Color.Blue)

                Dim bitLocker As ManagementClass = New ManagementClass(Me.ComputerContext.WMI.BitLockerScope, Me.ComputerContext.WMI.BitLockerScope.Path, New ObjectGetOptions)
                For Each bitLockerInstance As ManagementObject In bitLocker.GetInstances()

                    Dim protectionStatus(0) As Object
                    Dim hddEncryptionStatus(1) As Object

                    bitLockerInstance.InvokeMethod("GetProtectionStatus", protectionStatus)
                    bitLockerInstance.InvokeMethod("GetConversionStatus", hddEncryptionStatus)

                    If hddEncryptionStatus(0) = 1 Then
                        Select Case protectionStatus(0)

                            Case 0
                                TryWriteMessage(String.Format("BitLocker is disabled on the {0} drive. Prompting for consent to enable BitLocker", bitLockerInstance.Properties("DriveLetter").Value), Color.Blue)
                                If MsgBox(String.Format("Would you like to enable BitLocker for the {0} Drive?", bitLockerInstance.Properties("DriveLetter").Value), MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                                    bitLockerInstance.InvokeMethod("EnableKeyProtectors", Nothing)
                                End If

                                TryWriteMessage(String.Format("BitLocker has been successfully enabled on {0} drive", bitLockerInstance.Properties("DriveLetter").Value), Color.Blue)
                                If Me.ComputerContext IsNot Nothing Then
                                    Me.ComputerContext.SetConnectionStatus(ComputerControl.ConnectionStatuses.Online)
                                End If

                            Case 1
                                TryWriteMessage(String.Format("BitLocker is enabled on the {0} drive. Prompting for consent to disable BitLocker", bitLockerInstance.Properties("DriveLetter").Value), Color.Blue)
                                If MsgBox(String.Format("Would you like to disable BitLocker for the {0} Drive?", bitLockerInstance.Properties("DriveLetter").Value), MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                                    bitLockerInstance.InvokeMethod("DisableKeyProtectors", Nothing)
                                End If

                                TryWriteMessage(String.Format("BitLocker has been successfully disabled on {0} drive", bitLockerInstance.Properties("DriveLetter").Value), Color.Blue)
                                If Me.ComputerContext IsNot Nothing Then
                                    Me.ComputerContext.SetConnectionStatus(ComputerControl.ConnectionStatuses.Online)
                                End If

                        End Select
                    Else
                        TryWriteMessage(String.Format("No encryption detected on the {0} drive.", bitLockerInstance.Properties("DriveLetter").Value), Color.Blue)
                    End If
                Next
            Else
                TryWriteMessage("A connection has not been established with the BitLocker Namespace in WMI", Color.Red)
                TryWriteMessage("This is most likely because BitLocker is not installed on the computer", Color.Red)
            End If
            TryWriteMessage(String.Format("Terminated Toggle BitLocker on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub EnableComputers()
        Try
            TryWriteMessage("Connecting to computer account in LDAP", Color.Blue)

            Dim computer As New LDAPContainer("computer", "name", Me.ComputerContext.Computer.Value)
            If computer IsNot Nothing Then
                If Convert.ToBoolean(computer.GetAttribute("userAccountControl") And 2) Then
                    TryWriteMessage("Computer is disabled. Re-enabling...", Color.Blue)

                    computer.SetAttribute("userAccountControl", computer.GetAttribute("userAccountControl") - 2)

                    TryWriteMessage("Computer has been enabled!", Color.Blue)
                Else
                    TryWriteMessage("Computer is already enabled in LDAP. No action taken", Color.Blue)
                End If
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage(String.Format("Unable to enable computer account. {0}", ex.Message), Color.Red)
        End Try
    End Sub

    Private Sub DisableComputers()
        Try
            TryWriteMessage("Connecting to computer account in LDAP", Color.Blue)

            Dim computer As New LDAPContainer("computer", "name", Me.ComputerContext.Computer.Value)
            If computer IsNot Nothing Then
                If Not Convert.ToBoolean(computer.GetAttribute("userAccountControl") And 2) Then
                    TryWriteMessage("Computer is enabled. Disabling...", Color.Blue)

                    computer.SetAttribute("userAccountControl", computer.GetAttribute("userAccountControl") + 2)

                    TryWriteMessage("Computer has been disabled!", Color.Blue)
                Else
                    TryWriteMessage("Computer is already disabled in LDAP. No action taken", Color.Blue)
                End If
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage(String.Format("Unable to disable computer account. {0}", ex.Message), Color.Red)
        End Try
    End Sub

    Private Sub NoPromptRestart()
        Try
            TryWriteMessage(String.Format("Initializing Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            If Me.ComputerContext.WMI.BitLockerScope.IsConnected Then
                TryWriteMessage("BitLocker is installed on this computer. Checking status", Color.Blue)

                Dim bitLocker As ManagementClass = New ManagementClass(Me.ComputerContext.WMI.BitLockerScope, Me.ComputerContext.WMI.BitLockerScope.Path, New ObjectGetOptions)
                For Each bitLockerInstance As ManagementObject In bitLocker.GetInstances()
                    Dim protectionStatus(0) As Object
                    Dim hddEncryptionStatus(1) As Object

                    bitLockerInstance.InvokeMethod("GetProtectionStatus", protectionStatus)
                    bitLockerInstance.InvokeMethod("GetConversionStatus", hddEncryptionStatus)

                    If hddEncryptionStatus(0) = 1 AndAlso protectionStatus(0) = 1 Then
                        TryWriteMessage("BitLocker is enabled on this drive. Mass reboots automatically attempt to disable BitLocker", Color.Blue)

                        bitLockerInstance.InvokeMethod("DisableKeyProtectors", Nothing)
                        TryWriteMessage("BitLocker has been disabled", Color.Blue)
                    End If
                Next

                For Each win32_OperatingSystem As ManagementObject In Me.ComputerContext.WMI.Query("SELECT * FROM Win32_OperatingSystem WHERE Primary=TRUE")
                    win32_OperatingSystem.InvokeMethod("Reboot", Nothing)
                    TryWriteMessage(String.Format("{0} is currently rebooting", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
                Next
            End If

            TryWriteMessage(String.Format("Terminated Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub RemoteRestart()
        Try
            Dim instant As Boolean = True
            Dim delay As Integer = Nothing
            Dim userMessage As String = Nothing

            TryWriteMessage(String.Format("Initializing Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            If MsgBox(String.Format("Are you sure you want to restart{0}{1}?", Environment.NewLine, Me.ComputerContext.Computer.Display), MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then

                If Me.ComputerContext.WMI.BitLockerScope.IsConnected Then
                    TryWriteMessage("BitLocker is installed on this computer. Checking status", Color.Blue)

                    Dim bitLocker As ManagementClass = New ManagementClass(Me.ComputerContext.WMI.BitLockerScope, Me.ComputerContext.WMI.BitLockerScope.Path, New ObjectGetOptions())
                    For Each bitLockerInstance As ManagementObject In bitLocker.GetInstances()

                        Dim protectionStatus(0) As Object
                        Dim hddEncryptionStatus(1) As Object

                        bitLockerInstance.InvokeMethod("GetProtectionStatus", protectionStatus)
                        bitLockerInstance.InvokeMethod("GetConversionStatus", hddEncryptionStatus)

                        If hddEncryptionStatus(0) = 1 AndAlso protectionStatus(0) = 1 Then
                            TryWriteMessage("BitLocker is enabled on this drive. Prompting for consent to disable BitLocker", Color.Blue)

                            If MsgBox("Would you like to disable BitLocker before rebooting?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                                bitLockerInstance.InvokeMethod("DisableKeyProtectors", Nothing)
                                TryWriteMessage("BitLocker has been disabled", Color.Blue)
                            End If
                        End If
                    Next
                End If

                For Each win32_OperatingSystem As ManagementObject In Me.ComputerContext.WMI.Query("SELECT * FROM Win32_OperatingSystem WHERE Primary=TRUE")
                    win32_OperatingSystem.InvokeMethod("Reboot", Nothing)
                    TryWriteMessage(String.Format("{0} is currently rebooting", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
                Next
            End If

            TryWriteMessage(String.Format("Terminated Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub TimedRemoteRestart()
        Try
            Dim delay As Integer = Nothing
            Dim userMessage As String = Nothing

            TryWriteMessage(String.Format("Initializing Timed Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            If MsgBox(String.Format("Are you sure you want to restart{0}{1}?", Environment.NewLine, Me.ComputerContext.Computer.Display), MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then

                Dim userDelay As String = InputBox(String.Format("How long until the computer is rebooted (in seconds){0}Default: 30 Minutes (1800 Seconds)", Environment.NewLine), , "1800")
                If String.IsNullOrWhiteSpace(userDelay) OrElse Not Regex.IsMatch(userDelay, "[0-9]+") Then
                    TryWriteMessage("Invalid Time Format. Exiting Remote Restart", Color.Red)
                    TryWriteMessage(String.Format("Terminated Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
                    Exit Sub
                End If

                delay = Convert.ToInt32(userDelay)
                Dim userComment As String = InputBox("What message would you like displayed to the user? The time the restart will occur will be appended at the end of the message.", , String.Format("Your computer will be rebooted in {0} minutes", Math.Round(delay / 60, 0)))
                If String.IsNullOrWhiteSpace(userComment) Then
                    TryWriteMessage("Invalid Comment. Exiting Remote Restart", Color.Red)
                    TryWriteMessage(String.Format("Terminated Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
                    Exit Sub
                End If

                If Me.ComputerContext.WMI.BitLockerScope.IsConnected Then
                    TryWriteMessage("BitLocker is installed on this computer. Checking status", Color.Blue)

                    Dim bitLocker As ManagementClass = New ManagementClass(Me.ComputerContext.WMI.BitLockerScope, Me.ComputerContext.WMI.BitLockerScope.Path, New ObjectGetOptions)
                    For Each bitLockerInstance As ManagementObject In bitLocker.GetInstances()

                        Dim protectionStatus(0) As Object
                        Dim hddEncryptionStatus(1) As Object

                        bitLockerInstance.InvokeMethod("GetProtectionStatus", protectionStatus)
                        bitLockerInstance.InvokeMethod("GetConversionStatus", hddEncryptionStatus)

                        If hddEncryptionStatus(0) = 1 AndAlso protectionStatus(0) = 1 Then
                            TryWriteMessage("BitLocker is enabled on this drive. Prompting for consent to disable BitLocker", Color.Blue)
                            If MsgBox("Would you like to disable BitLocker before rebooting?", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                                bitLockerInstance.InvokeMethod("DisableKeyProtectors", Nothing)
                                TryWriteMessage("BitLocker has been disabled", Color.Blue)
                            End If
                        End If
                    Next
                End If


                For Each win32_OperatingSystem As ManagementObject In Me.ComputerContext.WMI.Query("SELECT * FROM Win32_OperatingSystem WHERE Primary=TRUE")
                    Dim shutdownTrackerParams As ManagementBaseObject = win32_OperatingSystem.GetMethodParameters("Win32ShutdownTracker")
                    With shutdownTrackerParams
                        .SetPropertyValue("Timeout", delay)
                        .SetPropertyValue("Comment", String.Format("{0}{1}Time of Restart: {2}", userComment, Environment.NewLine, Now.AddSeconds(delay).ToString("MM/dd/yyyy HH:mm:ss")))
                        .SetPropertyValue("ReasonCode", 218)
                        .SetPropertyValue("Flags", 6)
                    End With
                    win32_OperatingSystem.InvokeMethod("Win32ShutdownTracker", shutdownTrackerParams, Nothing)

                    TryWriteMessage(String.Format("{0} will reboot in {1} seconds", Me.ComputerContext.Computer.ConnectionString, delay), Color.Blue)
                Next
            End If

            TryWriteMessage(String.Format("Terminated Timed Remote Restart on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub RemoteLogoff()
        Try
            TryWriteMessage(String.Format("Initializing Remote Logoff on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)

            If Me.ComputerContext.WMI.Query("SELECT Name FROM Win32_Process WHERE Name='explorer.exe'", Me.ComputerContext.WMI.RegularScope).Count > 0 Then
                For Each win32_OperatingSystem As ManagementObject In Me.ComputerContext.WMI.Query("SELECT * FROM Win32_OperatingSystem WHERE Primary=TRUE")
                    Dim shutdownParams As ManagementBaseObject = win32_OperatingSystem.GetMethodParameters("Win32Shutdown")
                    With shutdownParams
                        .SetPropertyValue("Flags", 0)
                        .SetPropertyValue("Reserved", 0)
                    End With

                    win32_OperatingSystem.InvokeMethod("Win32Shutdown", shutdownParams, Nothing)
                Next
                TryWriteMessage("Remote Logoff has been performed", Color.Blue)
            Else
                TryWriteMessage("No user has been detected as logged into the computer", Color.Blue)
            End If

            TryWriteMessage(String.Format("Terminated Remote Logoff on {0}", Me.ComputerContext.Computer.ConnectionString), Color.Blue)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub RunCustomAction()
        Try
            TryWriteMessage("Attempting to locate PsExec.exe", Color.Blue)
            [Global].SetPsExecBinaryPath()

            If String.IsNullOrWhiteSpace(My.Settings.PsExecPath) OrElse Not File.Exists(My.Settings.PsExecPath) Then
                MsgBox("Unable to locate a valid copy of PsExec.exe")
                Exit Sub
            Else
                Dim customActionRegistryPath = Path.Combine(My.Settings.RegistryPathCustomActions, CustomAction)

                Dim customActions As String() = New RegistryController().GetKeyValues(customActionRegistryPath, RegistryHive.CurrentUser, methodName:="EnumValues")
                For Each action As String In customActions
                    TryWriteMessage(String.Format("Preparing Custom Action: {0}", CustomAction), Color.Blue)
                    Dim parsedCommand As String = New RegistryController().GetKeyValue(customActionRegistryPath, action, RegistryController.RegistryKeyValueTypes.String).ToUpper().Replace("|COMPUTER|", Me.ComputerContext.Computer.ConnectionString, RegistryHive.CurrentUser)

                    TryWriteMessage(String.Format("Running Custom Command: {0}", parsedCommand), Color.Blue)

                    Dim psExecProcess As New ProcessStartInfo() With
                    {
                        .UseShellExecute = False,
                        .FileName = "cmd.exe",
                        .Arguments = String.Format(" /k ""{0}"" -s \\{1} {2}", My.Settings.PsExecPath, Me.ComputerContext.Computer.ConnectionString, parsedCommand),
                        .Verb = "RunAs",
                        .WorkingDirectory = Environment.SystemDirectory
                    }

                    Process.Start(psExecProcess)
                Next
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage("An error occured while running the custom action", Color.Red)
        End Try
    End Sub

#End Region

#Region " Printer "

    Private Sub PrinterAdd()
        Try
            ' Install a printer on a remote computer
            Process.Start(String.Format("{0} /il /c \\{1}", Path.Combine(Environment.SystemDirectory, "Printui.exe"), Me.ComputerContext.Computer.ConnectionString)).WaitForExit()
            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PrinterAddCab()
        Try
            ' Install printers using a printerExport file on a remote computer
            Dim openDialog As New OpenFileDialog With
            {
                .CheckFileExists = True,
                .CheckPathExists = True,
                .Multiselect = True,
                .Title = "Select the Printer Export File(s) you would like to install",
                .Filter = "Printer Export Files|*.printerExport"
            }

            If openDialog.ShowDialog() = DialogResult.OK Then
                Dim randomNumber As New Random
                Dim openSharePrinterName As String = randomNumber.Next(9999999)

                TryWriteMessage("Creating a printer share", Color.Blue)

                Dim driverName As String = "Canon Inkjet iP100 series"
                Dim win32_Printer As New ManagementClass(Me.ComputerContext.WMI.RegularScope, New ManagementPath("Win32_Printer"), Nothing)
                Dim win32_PrinterInstance As ManagementObject = win32_Printer.CreateInstance()
                With win32_PrinterInstance
                    .Item("DriverName") = driverName
                    .Item("PortName") = "LPT1:"
                    .Item("DeviceID") = openSharePrinterName
                    .Item("Network") = False
                    .Item("Shared") = True
                    .Item("ShareName") = openSharePrinterName
                    .Item("PrintProcessor") = "WinPrint"
                    .Item("PrintJobDataType") = "RAW"
                    .Put()
                    .Dispose()
                End With
                win32_Printer.Dispose()

                For Each fileName As String In openDialog.FileNames
                    Dim cabFile As String = Path.Combine(Application.UserAppDataPath, String.Format("{0}.printerExport", randomNumber.Next(9999999)))
                    If File.Exists(cabFile) Then
                        File.Delete(cabFile)
                    End If

                    TryWriteMessage(String.Format("Copying export file to: {0}", cabFile), Color.Blue)
                    File.Copy(fileName, cabFile, True)
                    Dim printBrmProcess As New ProcessStartInfo With
                    {
                        .Verb = "RunAs",
                        .FileName = Path.Combine(Environment.SystemDirectory, "spool\tools\PrintBrm.exe"),
                        .Arguments = String.Format("-s {0} -r -f {1} -O FORCE", Me.ComputerContext.Computer.ConnectionString, cabFile),
                        .UseShellExecute = False
                    }

                    TryWriteMessage(String.Format("Running the export file: {0}", cabFile), Color.Blue)
                    Process.Start(printBrmProcess).WaitForExit()

                    TryWriteMessage(String.Format("Deleting the export file: {0}", cabFile), Color.Blue)
                    File.Delete(cabFile)
                Next

                TryWriteMessage("Deleting the Printer Share", Color.Blue)
                win32_PrinterInstance.Delete()

                RaiseEvent WorkCompleted(Me, Nothing)
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage("An error occured while attempting to load printers from an export. Please make sure to remove any printer share remnants.", Color.Red)
        End Try
    End Sub

    Private Sub PrinterNewCab()
        Try
            ' Create a backup of remote printers
            If Me.ComputerContext.WMI.Architecture = New WMIController(".", WMIController.ManagementScopes.Regular).Architecture Then
                Dim safeFileDialog As New SaveFileDialog() With
                    {
                        .CheckPathExists = True,
                        .Title = "Select the location to save the Printer Export file",
                        .Filter = "Printer Export Files|*.printerExport"
                    }

                If safeFileDialog.ShowDialog() = DialogResult.OK Then
                    Dim randomNumber As New Random
                    Dim openSharePrinterName As String = String.Format("DELETE THIS PRINTER - LEFTOVER FROM CAB CREATION ({0})", randomNumber.Next(9999999))
                    Dim cabFile As String = Path.Combine(Application.UserAppDataPath, String.Format("{0}.printerExport", randomNumber.Next(9999999)))

                    TryWriteMessage("Creating a printer share", Color.Blue)

                    Dim driverName As String = "Canon Inkjet iP100 series"

                    Dim printerPath As ManagementPath = Me.ComputerContext.WMI.RegularScope.Path
                    printerPath.ClassName = "Win32_Printer"

                    Dim win32_Printer As New ManagementClass(Me.ComputerContext.WMI.RegularScope, printerPath, Nothing)
                    Dim win32_PrinterInstance As ManagementObject = win32_Printer.CreateInstance()
                    With win32_PrinterInstance
                        .Item("DriverName") = driverName
                        .Item("PortName") = "LPT1:"
                        .Item("DeviceID") = openSharePrinterName
                        .Item("Network") = False
                        .Item("Shared") = True
                        .Item("ShareName") = openSharePrinterName
                        .Item("PrintProcessor") = "winprint"
                        .Item("PrintJobDataType") = "RAW"
                        .Put()
                        .Dispose()
                    End With
                    win32_Printer.Dispose()

                    If File.Exists(cabFile) Then File.Delete(cabFile)
                    Dim printBrmProcess As New ProcessStartInfo With
                    {
                        .Verb = "RunAs",
                        .FileName = Path.Combine(Environment.SystemDirectory, "spool\tools\PrintBrm.exe"),
                        .Arguments = String.Format("-s {0} -b -f {1} -O FORCE", Me.ComputerContext.Computer.ConnectionString, cabFile),
                        .UseShellExecute = False
                    }

                    TryWriteMessage("Running the backup process", Color.Blue)
                    Process.Start(printBrmProcess).WaitForExit()

                    TryWriteMessage("Deleting the Printer Share", Color.Blue)
                    win32_PrinterInstance.Delete()

                    File.Copy(cabFile, safeFileDialog.FileName)
                    File.Delete(cabFile)
                End If
            Else
                TryWriteMessage("The target computer must have the same processor architecture as this computer", Color.Red)
                TryWriteMessage(String.Format("This Computer: {0}", New WMIController(".", WMIController.ManagementScopes.Regular).Architecture.ToString()), Color.Red)
                TryWriteMessage(String.Format("Remote Computer: {0}", Me.ComputerContext.WMI.Architecture), Color.Red)
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            TryWriteMessage("An error occured while attempting to backup printers to a file. Please make sure to remove any printer share remnants.", Color.Red)
        End Try
    End Sub

    Private Sub PrinterRename()
        Try
            ' Rename a remote printer
            For Each printerWMIInstance As ManagementObject In WMIObjects
                Dim currentPrinterName As String = printerWMIInstance.Properties("Name").Value
                Dim newPrinterName(1) As String
                newPrinterName(0) = InputBox("What is the new Printer name:", , currentPrinterName)

                If newPrinterName IsNot Nothing AndAlso newPrinterName IsNot currentPrinterName Then
                    TryWriteMessage(String.Format("Renaming Printer from {0} to {1}", currentPrinterName, newPrinterName(0)), Color.Blue)
                    Me.ComputerContext.WMI.Query(String.Format("SELECT * FROM Win32_Printer WHERE Name=""{0}""", currentPrinterName))(0).InvokeMethod("RenamePrinter", newPrinterName)

                    TryWriteMessage(String.Format("Printer successfully renamed to {0}", newPrinterName(0)), Color.Blue)
                End If
            Next

            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PrinterSetDefault()
        Try
            ' Changes the default printer for the user logged into a remote computer
            If Me.ComputerContext.UserLoggedOn Then
                Dim curentUserRegistryKey As String = Nothing
                Dim x86Registry As New RegistryController(Me.ComputerContext.WMI.X86Scope)

                For Each allUsersKeyValue In x86Registry.GetKeyValues(String.Empty, RegistryHive.Users)
                    Select Case allUsersKeyValue
                        Case ".DEFAULT"
                            Continue For

                        Case "S-1-5-18"
                            Continue For

                        Case "S-1-5-19"
                            Continue For

                        Case "S-1-5-20"
                            Continue For

                        Case Else
                            For Each subKeyValue In x86Registry.GetKeyValues(allUsersKeyValue, RegistryHive.Users)
                                If subKeyValue = "Volatile Environment" Then
                                    curentUserRegistryKey = allUsersKeyValue
                                    Exit Select
                                End If
                            Next

                    End Select

                    If curentUserRegistryKey IsNot Nothing Then Exit For
                Next

                If curentUserRegistryKey Is Nothing Then
                    TryWriteMessage("Unable to locate the user's registry hive", Color.Red)
                Else
                    x86Registry.SetKeyValue(String.Format("{0}\Software\Microsoft\Windows NT\CurrentVersion\Windows", curentUserRegistryKey), "Device", String.Format("{0},winspool,Ne00:", WMIObjects(0).Properties("Name").Value), RegistryController.RegistryKeyValueTypes.String, RegistryHive.Users)
                End If

                RaiseEvent WorkCompleted(Me, Nothing)
            Else
                TryWriteMessage("A default printer can only be set when a user is logged in", Color.Red)
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PrinterOpenQueue()
        Try
            ' Checks the remote printer queue
            Dim rundll32Process As New ProcessStartInfo() With
            {
                .Verb = "RunAs",
                .FileName = Path.Combine(Environment.SystemDirectory, "rundll32.exe"),
                .Arguments = String.Format("printui.dll,PrintUIEntry /o /n ""\\{0}\{1}""", Me.ComputerContext.Computer.ConnectionString, Me.WMIObjects(0).Properties("Name").Value)
            }

            Process.Start(rundll32Process).WaitForExit()
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PrinterOpenProperties()
        Try
            ' Checks the properties of a remote printer
            Dim rundll32ProcessInfo As New ProcessStartInfo() With
            {
                .Verb = "RunAs",
                .FileName = Path.Combine(Environment.SystemDirectory, "rundll32.exe"),
                .Arguments = String.Format("printui.dll,PrintUIEntry /p /n ""\\{0}\{1}""", Me.ComputerContext.Computer.ConnectionString, Me.WMIObjects(0).Properties("Name").Value)
            }

            Process.Start(rundll32ProcessInfo).WaitForExit()
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PrinterEditPort()
        Try
            ' Change the TCP/IP Printer port of a printer installed on a remote computer
            For Each printerWMIObject As ManagementObject In Me.WMIObjects
                Dim portName As String = InputBox(Prompt:="Enter either an existing port, or a new port:", DefaultResponse:=printerWMIObject.Properties("PortName").Value)
                Dim fullPrinterObject As ManagementObject = Me.ComputerContext.WMI.Query(String.Format("SELECT * FROM Win32_Printer WHERE Name=""{0}""", printerWMIObject.Properties("Name").Value))(0)
                Dim oldTCPIPPort As ManagementObject = Me.ComputerContext.WMI.Query(String.Format("SELECT Protocol, PortNumber, SNMPEnabled FROM Win32_TCPIPPrinterPort WHERE Name=""{0}""", printerWMIObject.Properties("PortName").Value))(0)

                If portName IsNot Nothing AndAlso portName IsNot printerWMIObject.Properties("PortName").Value Then
                    Dim existingPort As Boolean = False
                    For Each tcpIPPort As ManagementObject In Me.ComputerContext.WMI.Query(String.Format("SELECT * FROM Win32_TCPIPPrinterPort Where Name LIKE ""{0}""%", portName))
                        If tcpIPPort.Properties("HostAddress").Value = portName OrElse tcpIPPort.Properties("PortName").Value.ToString().StartsWith(String.Format("{0}_", portName)) Then
                            existingPort = True

                            TryWriteMessage(String.Format("Existing Port {0} found! Setting Port", tcpIPPort.Properties("Name").Value), Color.Blue)
                            fullPrinterObject("PortName") = portName

                            fullPrinterObject.Put()
                            TryWriteMessage("Port Successfully Changed", Color.Blue)
                        End If
                    Next

                    If Not existingPort Then
                        TryWriteMessage(String.Format("Creating new TCP/IP Port: {0}", portName), Color.Blue)
                        Dim win32_TCPIPPrinterPort As New ManagementClass(Me.ComputerContext.WMI.RegularScope, New ManagementPath("Win32_TCPIPPrinterPort"), Nothing)

                        Dim win32_TCPIPPrinterPortInstance As ManagementObject = win32_TCPIPPrinterPort.CreateInstance()
                        With win32_TCPIPPrinterPortInstance
                            .Properties("Name").Value = portName
                            .Properties("Protocol").Value = oldTCPIPPort.Properties("Protocol").Value
                            .Properties("HostAddress").Value = portName
                            .Properties("PortNumber").Value = oldTCPIPPort.Properties("PortNumber").Value
                            .Properties("SNMPEnabled").Value = oldTCPIPPort.Properties("SNMPEnabled").Value
                            .Put()
                            .Dispose()
                        End With
                        win32_TCPIPPrinterPort.Dispose()

                        fullPrinterObject("PortName") = portName
                        fullPrinterObject.Put()

                        TryWriteMessage("Port Successfully Changed", Color.Blue)
                    End If
                End If
            Next

            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub PrinterDeleteAnyType()
        Try
            ' Delete Printers, Printer Ports, and Printer Drivers installed on a remote computer
            For Each printerWMIObject As ManagementObject In Me.WMIObjects
                Dim printerName As String = printerWMIObject.Properties("Name").Value
                Dim displayText As String = Nothing

                Select Case printerWMIObject.Properties("CreationClassName").Value
                    Case "Win32_Printer"
                        displayText = "Printer Deletion"

                    Case "Win32_PrinterDriver"
                        displayText = "Print Driver Deletion"

                    Case "Win32_TCPIPPrinterPort"
                        displayText = "TCP/IP Printer Port Deletion"

                End Select

                TryWriteMessage(String.Format("Initializing {0} for {1} on {2}", displayText, printerName, Me.ComputerContext.Computer.ConnectionString), Color.Blue)

                Try
                    Me.ComputerContext.WMI.Query(String.Format("SELECT * FROM {0} WHERE Name = ""{1}""", printerWMIObject.Properties("CreationClassName").Value, printerName))(0).Delete()
                    TryWriteMessage(String.Format("{0} has been successfully deleted.", printerName), Color.Blue)
                Catch ex As Exception
                    LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                    TryWriteMessage("An error occured while attempting the deletion", Color.Red)
                Finally
                    printerWMIObject.Dispose()
                End Try
            Next

            RaiseEvent WorkCompleted(Me, EventArgs.Empty)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

#End Region

#Region " Service "

    Private Sub ServiceStop()
        Try
            For Each serviceWMIObject As ManagementObject In Me.WMIObjects
                TryWriteMessage(String.Format("Sending a stop request to {0}", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)

                Dim service As New ServiceController(Me.ComputerContext.WMI)
                Dim serviceStopResult As ServiceController.ServiceError = service.Stop(serviceWMIObject.Properties("Name").Value)

                If serviceStopResult = ServiceController.ServiceError.DependentServicesRunning Then
                    TryWriteMessage("The follwing dependent services are running and must be stopped first:", Color.Red)

                    For Each dependentService As String In service.CheckDependentServices(serviceWMIObject.Properties("Name").Value)
                        Dim win32_Service As ManagementObject = Me.ComputerContext.WMI.Query(String.Format("SELECT DisplayName, State FROM Win32_Service WHERE Name=""{0}""", dependentService))(0)

                        If win32_Service.Properties("State").Value.ToString().ToUpper().Trim() = "RUNNING" Then
                            TryWriteMessage(win32_Service.Properties("DisplayName").Value, Color.Red)
                        End If
                    Next
                Else
                    If service.WaitForService(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceState.Stopped, 10) Then
                        TryWriteMessage(String.Format("The service ""{0}"" has been stopped successfully", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)
                    Else
                        TryWriteMessage(String.Format("A stop request was sent to ""{0}"", but it did not stop in a timely manner", serviceWMIObject.Properties("DisplayName").Value), Color.Red)
                    End If

                    RaiseEvent WorkCompleted(Me, Nothing)
                End If
            Next
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ServiceStart()
        Try
            For Each serviceWMIObject As ManagementObject In Me.WMIObjects
                TryWriteMessage(String.Format("Sending a start request to {0}", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)

                Dim service As New ServiceController(Me.ComputerContext.WMI)
                service.Start(serviceWMIObject.Properties("Name").Value)

                If service.WaitForService(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceState.Running, 10) Then
                    TryWriteMessage(String.Format("The service ""{0}"" has been started successfully", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)
                Else
                    TryWriteMessage(String.Format("A start request was sent to ""{0}"", but it did not start in a timely manner", serviceWMIObject.Properties("DisplayName").Value), Color.Red)
                End If

                RaiseEvent WorkCompleted(Me, Nothing)
            Next
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ServiceRestart()
        Try
            For Each serviceWMIObject As ManagementObject In Me.WMIObjects
                TryWriteMessage(String.Format("Sending a stop request to {0}", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)

                Dim service As New ServiceController(Me.ComputerContext.WMI)
                Dim serviceStopResult As ServiceController.ServiceError = service.Stop(serviceWMIObject.Properties("Name").Value)

                If serviceStopResult = ServiceController.ServiceError.DependentServicesRunning Then
                    TryWriteMessage("The follwing dependent services are running and must be stopped first:", Color.Red)

                    For Each dependentService As String In service.CheckDependentServices(serviceWMIObject.Properties("Name").Value)
                        Dim win32_Service As ManagementObject = Me.ComputerContext.WMI.Query(String.Format("SELECT DisplayName, State FROM Win32_Service WHERE Name=""{0}""", dependentService))(0)
                        If win32_Service.Properties("State").Value.ToString().ToUpper().Trim() = "RUNNING" Then
                            TryWriteMessage(win32_Service.Properties("DisplayName").Value, Color.Red)
                        End If
                    Next
                Else
                    If service.WaitForService(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceState.Stopped, 10) Then
                        TryWriteMessage(String.Format("The service ""{0}"" has been stopped successfully", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)
                        TryWriteMessage(String.Format("Sending a start request to {0}", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)
                        service.Start(serviceWMIObject.Properties("Name").Value)

                        If service.WaitForService(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceState.Running, 10) Then
                            TryWriteMessage(String.Format("The service ""{0}"" has been started successfully", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)
                        Else
                            TryWriteMessage(String.Format("A start request was sent to ""{0}"", but it did not start in a timely manner", serviceWMIObject.Properties("DisplayName").Value), Color.Red)
                        End If
                    Else
                        TryWriteMessage(String.Format("A stop request was sent to ""{0}"", but it did not stop in a timely manner", serviceWMIObject.Properties("DisplayName").Value), Color.Red)
                    End If

                    RaiseEvent WorkCompleted(Me, Nothing)
                End If
            Next
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ServiceSetAuto()
        Try
            For Each serviceWMIObject As ManagementObject In Me.WMIObjects
                TryWriteMessage(String.Format("Changing Start Mode for {0} to Automatic", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)

                Dim service As New ServiceController(Me.ComputerContext.WMI)
                service.ChangeStartupType(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceStartupType.Auto)
            Next

            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ServiceSetManual()
        Try
            For Each serviceWMIObject As ManagementObject In Me.WMIObjects
                TryWriteMessage(String.Format("Changing Start Mode for {0} to Manual", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)

                Dim service As New ServiceController(Me.ComputerContext.WMI)
                service.ChangeStartupType(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceStartupType.Manual)
            Next

            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ServiceSetDisabled()
        Try
            For Each serviceWMIObject As ManagementObject In Me.WMIObjects
                TryWriteMessage(String.Format("Changing Start Mode for {0} to Disabled", serviceWMIObject.Properties("DisplayName").Value), Color.Blue)

                Dim service As New ServiceController(Me.ComputerContext.WMI)
                service.ChangeStartupType(serviceWMIObject.Properties("Name").Value, ServiceController.ServiceStartupType.Disabled)
            Next

            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

#End Region

#Region " Process "

    Private Sub ProcessStop()
        Try
            Dim process As New ProcessController(Me.ComputerContext.WMI)

            For Each processWMIObject As ManagementObject In WMIObjects
                TryWriteMessage(String.Format("Sending a stop request to ""{0} ({1})""", processWMIObject.Properties("Name").Value, processWMIObject.Properties("ProcessID").Value), Color.Blue)

                Dim processStopResult As ProcessController.ProcessError = process.StopProcess(processWMIObject.Properties("ProcessID").Value)
                If process.WaitForProcess(processWMIObject.Properties("Name").Value, ProcessController.ProcessCondition.Stopped, 10) AndAlso processStopResult = ProcessController.ProcessError.Success Then
                    TryWriteMessage(String.Format("The process ""{0} ({1})"" has been stopped successfully", processWMIObject.Properties("Name").Value, processWMIObject.Properties("ProcessID").Value), Color.Blue)
                Else
                    TryWriteMessage(String.Format("A stop request was sent to ""{0} ({1})"", but it did not stop in a timely manner", processWMIObject.Properties("Name").Value, processWMIObject.Properties("ProcessID").Value), Color.Red)
                End If
            Next

            RaiseEvent WorkCompleted(Me, Nothing)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

#End Region

#End Region

#Region " Notifications "

    ''' <summary>
    ''' Will only work if the remote tools are initialized with the sender passed to it, and the send has a "Write" subroutine that will display messages. This is currently only used by the ComputerControl object to display messages to the user
    ''' </summary>
    ''' <param name="message">The message to pass</param>
    ''' <param name="color">The color of the message</param>
    ''' <remarks></remarks>
    Private Sub TryWriteMessage(message As String, color As Color)
        If Me.ComputerContext IsNot Nothing Then
            Me.ComputerContext.WriteMessage(message, color)
        End If
    End Sub

    ''' <summary>
    ''' Will only work if the remote tools are initialized with the sender passed to it, and the send has a "Write" subroutine that will display messages. This is currently only used by the ComputerControl object to display messages to the user
    ''' </summary>
    ''' <param name="message">The message to pass</param>
    ''' <remarks></remarks>
    Private Sub TryWriteMessage(message As String)
        TryWriteMessage(message, Color.Black)
    End Sub

#End Region

End Class