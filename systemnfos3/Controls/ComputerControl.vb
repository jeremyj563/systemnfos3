Imports System.ComponentModel
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports Microsoft.Win32
Imports systemnfos3.WMIController

Public Class ComputerControl
    Public Property OwnerForm As MainForm
    Public Property OwnerNode As TreeNode
    Public Property IsLoaded As Boolean
    Public Property UserStatus As UserStatuses = UserStatuses.None
    Public Property ConnectionStatus As ConnectionStatuses = ConnectionStatuses.Offline
    Public Property Computer As Computer
    Public Property WMI As WMIController
    Public Property Initialized As Boolean = False
    Public Property Timeout As Integer = 20
    Public Property IsMassImport As Boolean = False
    Public Property UseRandomTimeoutForMassImport As Boolean = False
    Public Property LoaderBackgroundThread As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
    Public Property PingLock As New Object
    Public Property UserLoggedOn As Boolean = False

    Private Property ForceFullOnlineMode As Boolean = False
    Private Property RebootPending As Boolean = False
    Private Property ResolveIPAddress As Boolean = False

    Friend Property LastSelectedTab As Tab = Nothing

    Public Sub New(computer As Computer, ownerForm As MainForm, ownerNode As TreeNode)
        ' This call is required by the designer.
        InitializeComponent()

        Me.Computer = computer
        Me.OwnerForm = ownerForm
        Me.OwnerNode = ownerNode

        Me.IsLoaded = False

        AddHandler Me.LoaderBackgroundThread.DoWork, AddressOf LoadComputer
    End Sub

#Region " Misc "

    Public Sub WriteMessage(message As String, Optional color As Color = Nothing, Optional isMassImport As Boolean = False)
        If isMassImport Then Exit Sub
        If color = Nothing Then color = Color.Black

        If Me.IsHandleCreated Then
            Me.UIThread(Sub()
                            With Me.StatusRichTextBox
                                .SelectionStart = .Text.Length
                                .SelectionColor = color
                                .AppendText(String.Format("({0}): {1}{2}", Now, message, Environment.NewLine))
                                .SelectionColor = .ForeColor
                                .SelectionStart = .Text.Length
                                .ScrollToCaret()
                            End With
                        End Sub)
        End If
    End Sub

    Public Sub ReportConnectionStatus()
        SyncLock PingLock
            Me.Timeout = 10
            Me.IsMassImport = False

            Select Case Me.ConnectionStatus

                Case ConnectionStatuses.Offline
                    If RespondsToPing() Then
                        WriteMessage("Computer has come online", Color.Green)
                        StartLoadingComputer()
                    End If

                Case Else
                    If Not RespondsToPing() Then
                        WriteMessage("Computer has gone offline", Color.Red)
                        SetConnectionStatus(ConnectionStatuses.Offline)
                    Else
                        If ConnectionStatus = ConnectionStatuses.Online Then
                            SetRebootPendingFont()
                            Dim userStatus As UserStatuses = GetUserStatus()
                            If Me.UserStatus <> userStatus Then
                                Me.UserStatus = userStatus
                                SetConnectionStatusColors(ConnectionStatuses.Online)

                                If TypeOf LastSelectedTab Is MainTab AndAlso Not LastSelectedTab.LoaderBackgroundThread.IsBusy Then
                                    LastSelectedTab.LoaderBackgroundThread.RunWorkerAsync()
                                End If
                            End If
                        Else
                            If IsConnectionFastEnoughForFullOnline() Then
                                WriteMessage("Computer connection speed is excellent")
                                StartLoadingComputer()
                            End If
                        End If
                    End If

            End Select
        End SyncLock
    End Sub

    Private Sub StatusRichTextBox_LinkClicked(sender As Object, e As LinkClickedEventArgs) Handles StatusRichTextBox.LinkClicked
        Try
            Process.Start(e.LinkText)
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub AddToCollection(sender As Object, e As EventArgs)
        Dim collectionName As String = CType(sender, ToolStripItem).Text
        CollectionController.AddToCollection(collectionName, Me.Computer.Value)
    End Sub

    Private Sub RemoveOwnerNode()
        SyncLock Me.PingLock
            Dim garbageCollector As New GarbageCollector With {.AddToGarbage = Me}

            Me.UIThread(Sub() Me.OwnerNode.Remove())

            Me.OwnerForm.UserSettingsRemoveComputer(Me.Computer)
            Me.OwnerForm.RefreshComputerListView()
            garbageCollector.DisposeAsync()
        End SyncLock
    End Sub

    Private Sub UnloadHandle()
        If Me.IsHandleCreated Then
            Me.DestroyHandle()
            Me.Initialized = False
        End If
    End Sub

#End Region

#Region " Remote Computer Connection Initiation "

    Private Sub LoadFullOnline()
        Me.ForceFullOnlineMode = True

        If Not Me.LoaderBackgroundThread.IsBusy Then
            Me.LoaderBackgroundThread.RunWorkerAsync()
        End If
    End Sub

    Private Sub EstablishWMIConnection(sender As BackgroundWorker, e As DoWorkEventArgs)
        Try
            Me.WMI = New WMIController(Me.Computer.ConnectionString)
            e.Result = Nothing
        Catch ex As Exception
            e.Result = ex
        End Try
    End Sub

    Public Sub LoadComputer(sender As Object, e As DoWorkEventArgs)
        StartLoadingComputer()
    End Sub

    Public Sub StartLoadingComputer()
        Me.Initialized = True

        If Not Me.IsMassImport Then
            WriteMessage(String.Format("Checking Connection Status of: {0}", Computer.Value))
        End If

        If RespondsToPing() Then
            WriteMessage("Computer responded to ping request", Color.Green, Me.IsMassImport)
            If Not Me.ForceFullOnlineMode Then
                WriteMessage("Testing Connection Speed", Color.Green, Me.IsMassImport)
            End If

            If IsConnectionFastEnoughForFullOnline(Me.IsMassImport) Then
                If Not Me.ForceFullOnlineMode Then
                    WriteMessage("Connection speed is excellent", Color.Green, Me.IsMassImport)
                End If

                WriteMessage("Initiating remote WMI connection", Color.Black, Me.IsMassImport)

                If Me.Timeout > 0 Then
                    Dim wmiWorker As New BackgroundWorker With {.WorkerSupportsCancellation = True, .WorkerReportsProgress = True}
                    Dim connectionException As Exception = Nothing

                    AddHandler wmiWorker.DoWork, AddressOf EstablishWMIConnection
                    AddHandler wmiWorker.RunWorkerCompleted, Sub(s As Object, ea As RunWorkerCompletedEventArgs) If ea.Result IsNot Nothing Then connectionException = ea.Result

                    Try
                        wmiWorker.RunWorkerAsync(True)
                    Catch ex As Exception
                        LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                        WriteMessage("An error occured connecting to the remote computer", Color.Red, Me.IsMassImport)
                        WriteMessage(ex.Message, Color.Red, Me.IsMassImport)
                    End Try

                    Dim i As Integer = 0
                    Do Until i = Me.Timeout
                        If connectionException IsNot Nothing Then
                            WriteMessage("An error occured connecting to the remote computer", Color.Red, Me.IsMassImport)
                            WriteMessage(connectionException.Message, Color.Brown, Me.IsMassImport)

                            SetConnectionStatus(ConnectionStatuses.OnlineDegraded)
                            Exit Sub
                        End If

                        For j = 0 To 10
                            If Me.WMI Is Nothing Then
                                Threading.Thread.Sleep(100)
                            Else
                                Exit Do
                            End If
                        Next
                        i += 1
                    Loop

                    wmiWorker.Dispose()

                    If Me.WMI Is Nothing Then
                        Dim message = String.Format("The WMI connection has been terminated because the timeout period of {0} seconds has elapsed", Me.Timeout)
                        WriteMessage(message, Color.Brown, Me.IsMassImport)
                        SetConnectionStatus(ConnectionStatuses.OnlineDegraded)
                        Exit Sub
                    End If
                Else
                    Try
                        Me.WMI = New WMIController(Computer.ConnectionString)
                    Catch ex As Exception
                        LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                        WriteMessage("An error occured while connecting to the remote computer", Color.Brown, IsMassImport)
                        WriteMessage(ex.Message, Color.Red, IsMassImport)
                        SetConnectionStatus(ConnectionStatuses.OnlineDegraded)
                        Exit Sub
                    End Try
                End If

                If Me.WMI.RegularScope.IsConnected Then
                    WriteMessage("Main WMI namespace connection successful!", Color.Green)
                Else
                    WriteMessage("Main WMI namespace connection unsuccessful!", Color.Red)
                End If

                If Me.WMI.LDAPScope.IsConnected Then
                    WriteMessage("LDAP WMI namespace connection successful!", Color.Green)
                Else
                    WriteMessage("LDAP WMI namespace connection unsuccessful!", Color.Red)
                End If

                If Me.WMI.X86Scope.IsConnected Then
                    WriteMessage("32-bit Registry WMI namespace connection successful!", Color.Green)
                Else
                    WriteMessage("32-bit Registry WMI namespace connection unsuccessful!", Color.Red)
                End If

                If Me.WMI.X64Scope.IsConnected Then
                    WriteMessage("64-bit Registry WMI namespace connection successful!", Color.Green)
                Else
                    WriteMessage("64-bit Registry WMI namespace connection unsuccessful!", Color.Red)
                End If

                If Me.WMI.BitLockerScope.IsConnected Then
                    WriteMessage("BitLocker WMI namespace connection successful!", Color.Green)
                Else
                    WriteMessage("BitLocker WMI namespace connection unsuccessful!", Color.Red)
                End If

                If Me.WMI.TPMScope.IsConnected Then
                    WriteMessage("TPM WMI namespace connection successful!", Color.Green)
                Else
                    WriteMessage("TPM WMI namespace connection unsuccessful!", Color.Red)
                End If

                SetConnectionStatus(ConnectionStatuses.Online)
            Else
                WriteMessage("Connection speed is poor", Color.Magenta)
                SetConnectionStatus(ConnectionStatuses.OnlineSlow)
            End If
        Else
            WriteMessage("Computer is not responding to ping request", Color.Red)
            SetConnectionStatus(ConnectionStatuses.Offline)
        End If

        Me.IsLoaded = True
    End Sub

    Public Function RespondsToPing() As Boolean
        Try
            If Not Me.LoaderBackgroundThread.CancellationPending AndAlso Not String.IsNullOrWhiteSpace(Me.Computer.Value) Then
                Dim ipAddress As IPAddress = Dns.GetHostAddresses(Me.Computer.Value).FirstOrDefault()
                If ipAddress IsNot Nothing Then
                    Dim ping As New Ping()

                    If ping.Send(ipAddress, 500).Status = IPStatus.Success Then
                        If Me.WMI Is Nothing Then
                            Me.Computer.ConnectionString = Me.Computer.Value
                        End If

                        Return True
                    Else
                        ' Try ping twice before giving up
                        If ping.Send(ipAddress, 500).Status = IPStatus.Success Then
                            Dim wmi As New WMIController(Me.Computer.IPAddress, ManagementScopes.Regular)
                            If wmi.GetPropertyValue(wmi.Query("SELECT Name FROM Win32_ComputerSystem"), "Name").ToUpper() = Me.Computer.Value.ToUpper() Then
                                Me.Computer.ConnectionString = Me.Computer.Value
                                Return True
                            Else
                                Me.Computer.ConnectionString = Me.Computer.IPAddress
                            End If
                        End If
                    End If
                End If
            End If
        Catch ex As Exception
            ' Silently fail
        End Try

        Return False
    End Function

    Private Function IsConnectionFastEnoughForFullOnline(Optional isMassImport As Boolean = False, Optional connectionString As String = Nothing, Optional attempt As Integer = 1) As Boolean
        If connectionString = Nothing Then connectionString = Computer.ConnectionString

        Dim wmi = New WMIController(".", ManagementScopes.Regular)
        Dim responseTime As String = wmi.GetPropertyValue(wmi.Query(String.Format("SELECT ResponseTime FROM Win32_PingStatus Where Address='{0}'", connectionString)), "ResponseTime")
        Dim responseTimeThreshold As Integer = If(isMassImport, 50, 10)

        If Not String.IsNullOrWhiteSpace(responseTime) AndAlso Integer.Parse(responseTime) < responseTimeThreshold Then
            Return True
        Else
            If attempt = 1 Then
                Return IsConnectionFastEnoughForFullOnline(isMassImport, connectionString, attempt + 1)
            End If
        End If

        Return Me.ForceFullOnlineMode
    End Function

#End Region

#Region " Context Menu Handlers "

    Private Sub AddMenuStrip(menuStrip As ContextMenuStrip, container As Object)
        Me.UIThread(Sub() container.ContextMenuStrip = menuStrip)
    End Sub

    Private Sub PopulateDynamicMenuStripEntries(sender As Object, e As EventArgs)
        Dim menuStrip As ContextMenuStrip = sender
        If CollectionController.GetAllCollections IsNot Nothing AndAlso CollectionController.GetAllCollections.Count > 0 Then

            Dim itemExists As Boolean = False
            For Each item As ToolStripItem In menuStrip.Items
                If item.Text.Trim().ToUpper() = "ADD TO COLLECTION" Then
                    itemExists = True
                    Dim menuItem As ToolStripMenuItem = item
                    menuItem.DropDownItems.Clear()
                    CollectionController.GetAllCollections.ToList().ForEach(Function(collection) menuItem.DropDownItems.Add(collection))
                    Exit For
                End If
            Next

            If Not itemExists Then
                Dim menuItem As New ToolStripMenuItem("Add to Collection")
                CollectionController.GetAllCollections.ToList().ForEach(Function(collection) menuItem.DropDownItems.Add(collection, Nothing, AddressOf AddToCollection))
            End If
        Else
            For i As Integer = 0 To menuStrip.Items.Count - 1
                If menuStrip.Items(i).Text.Trim().ToUpper() = "ADD TO COLLECTION" Then
                    menuStrip.Items.Remove(menuStrip.Items(i - 1))
                    menuStrip.Items.Remove(menuStrip.Items(i))
                    Exit For
                End If
            Next
        End If

        ' Remove Custom Actions from Right click
        Dim customActionExists As Boolean = False
        For index = menuStrip.Items.Count - 1 To 0 Step -1
            If menuStrip.Items(index).Tag IsNot Nothing Then
                If TypeOf menuStrip.Items(index).Tag Is String() Then
                    Dim tag As String() = menuStrip.Items(index).Tag
                    If tag.Count > 1 Then
                        If tag(1).ToString().ToUpper().Trim() = "CUSTOM" Then
                            customActionExists = True
                            menuStrip.Items.RemoveAt(index)
                        End If
                    End If
                End If
            End If
        Next

        ' Remove the Tool Strip Seperator if custom actions used to exist
        If customActionExists Then
            menuStrip.Items.RemoveAt(menuStrip.Items.Count - 1)
        End If

        If ConnectionStatus = ConnectionStatuses.Online OrElse ConnectionStatus = ConnectionStatuses.OnlineDegraded OrElse ConnectionStatus = ConnectionStatuses.OnlineSlow Then
            ' Repopulate Custom actions
            For Each action In New RegistryController().GetKeyValues(My.Settings.RegistryPathCustomActions, RegistryHive.CurrentUser)
                Dim menuItem As New ToolStripMenuItem(action, Nothing, AddressOf InitiateCustomAction)
                menuItem.Tag = New String() {action, "CUSTOM"}
                menuStrip.Items.Add(menuItem)
            Next
        End If
    End Sub

    Public Function NewComputerMenuStrip(status As ConnectionStatuses) As ContextMenuStrip
        Dim menuStrip As New ContextMenuStrip()
        With menuStrip

            .Items.Add("Remove From List", Nothing, AddressOf RemoveOwnerNode)
            Select Case status
                Case ConnectionStatuses.Offline
                    .Items.Add(New ToolStripSeparator())
                    .Items.Add("Set Description", Nothing, AddressOf InitiateSetDescription)
                    .Items.Add("Set Location", Nothing, AddressOf InitiateSetLocation)
                    Dim changeADStatusMenuItem As New ToolStripMenuItem("Change AD Status")
                    changeADStatusMenuItem.DropDownItems.Add(New ToolStripMenuItem("Enable Computer", Nothing, AddressOf InitiateEnableComputers))
                    changeADStatusMenuItem.DropDownItems.Add(New ToolStripMenuItem("Disable Computer", Nothing, AddressOf InitiateDisableComputers))
                    .Items.Add(changeADStatusMenuItem)

                Case ConnectionStatuses.OnlineDegraded
                    .Items.Add(New ToolStripSeparator())
                    Dim cDriveShareMenuItem As New ToolStripMenuItem("C$", Nothing, AddressOf InitiateAdminShare)
                    cDriveShareMenuItem.Tag = "C$"
                    .Items.Add(cDriveShareMenuItem)
                    .Items.Add("PsExec", Nothing, AddressOf InitiatePsExec)
                    .Items.Add("Computer Management", Nothing, AddressOf InitiateComputerManagement)
                    .Items.Add("Set Description", Nothing, AddressOf InitiateSetDescription)
                    .Items.Add("Set Location", Nothing, AddressOf InitiateSetLocation)
                    Dim changeADStatusMenuItem As New ToolStripMenuItem("Change AD Status")
                    changeADStatusMenuItem.DropDownItems.Add(New ToolStripMenuItem("Enable Computer", Nothing, AddressOf InitiateEnableComputers))
                    changeADStatusMenuItem.DropDownItems.Add(New ToolStripMenuItem("Disable Computer", Nothing, AddressOf InitiateDisableComputers))
                    .Items.Add(changeADStatusMenuItem)
                    .Items.Add(New ToolStripSeparator())

                Case ConnectionStatuses.Online, ConnectionStatuses.OnlineSlow
                    If status = ConnectionStatuses.OnlineSlow Then .Items.Add(New ToolStripSeparator())
                    If status = ConnectionStatuses.OnlineSlow Then .Items.Add("Load Online Information", Nothing, AddressOf LoadFullOnline)
                    .Items.Add(New ToolStripSeparator())
                    .Items.Add("Remote Assistance", Nothing, AddressOf InitiateRemoteAssistance)
                    .Items.Add("Remote Desktop", Nothing, AddressOf InitiateRemoteDesktop)
                    .Items.Add("Remote Control Viewer", Nothing, AddressOf InitiateRemoteControlViewer)
                    .Items.Add("Remote Registry", Nothing, AddressOf InitiateRemoteRegistry)
                    Dim hardDriveSharesMenuItem As New ToolStripMenuItem("Hard Drive Share(s)")
                    If Me.WMI IsNot Nothing Then
                        Dim wmiQueryResult = Me.WMI.Query("SELECT Name, Path FROM Win32_Share WHERE Description=""Default Share""")
                        If wmiQueryResult IsNot Nothing Then
                            For Each share As ManagementBaseObject In wmiQueryResult
                                Dim shareMenuItem As New ToolStripMenuItem(share.Properties("Path").Value, Nothing, AddressOf InitiateAdminShare)
                                shareMenuItem.Tag = share.Properties("Name").Value
                                hardDriveSharesMenuItem.DropDownItems.Add(shareMenuItem)
                            Next
                            If hardDriveSharesMenuItem.DropDownItems.Count > 0 Then .Items.Add(hardDriveSharesMenuItem)
                        End If
                    End If
                    .Items.Add("PsExec", Nothing, AddressOf InitiatePsExec)
                    .Items.Add("Computer Management", Nothing, AddressOf InitiateComputerManagement)
                    .Items.Add("Remote Group Policy Editor", Nothing, AddressOf InitiateGroupPolicyEditor)
                    .Items.Add("Set Description", Nothing, AddressOf InitiateSetDescription)
                    .Items.Add("Set Location", Nothing, AddressOf InitiateSetLocation)
                    Dim changeADStatusMenuItem As New ToolStripMenuItem("Change AD Status")
                    changeADStatusMenuItem.DropDownItems.Add(New ToolStripMenuItem("Enable Computer", Nothing, AddressOf InitiateEnableComputers))
                    changeADStatusMenuItem.DropDownItems.Add(New ToolStripMenuItem("Disable Computer", Nothing, AddressOf InitiateDisableComputers))
                    .Items.Add(changeADStatusMenuItem)
                    .Items.Add("Toggle BitLocker", Nothing, AddressOf InitiateToggleBitLocker)
                    .Items.Add("Remote Restart", Nothing, AddressOf InitiateRestart)
                    .Items.Add("Remote Logoff", Nothing, AddressOf InitiateLogoff)

            End Select

            AddHandler .Opening, AddressOf PopulateDynamicMenuStripEntries
        End With

        Return menuStrip
    End Function

#End Region

#Region " Online Mode Initiation "

    Private Function GetUserStatus() As UserStatuses
        If ConnectionStatus = ConnectionStatuses.Online AndAlso Me.WMI IsNot Nothing Then
            Dim wmiQueryResult As String = Me.WMI.GetPropertyValue(Me.WMI.Query("SELECT Name FROM Win32_Process WHERE Name=""explorer.exe"""), "Name")
            If Not String.IsNullOrWhiteSpace(wmiQueryResult) AndAlso wmiQueryResult.ToUpper() = "EXPLORER.EXE" Then
                wmiQueryResult = Me.WMI.GetPropertyValue(Me.WMI.Query("SELECT Name FROM Win32_Process WHERE Name=""logonui.exe"""), "Name")
                If Not String.IsNullOrWhiteSpace(wmiQueryResult) AndAlso wmiQueryResult.ToUpper() = "LOGONUI.EXE" Then
                    Return UserStatuses.Inactive
                End If
            Else
                Return UserStatuses.None
            End If
        End If

        Return UserStatuses.Active
    End Function

    Public Enum UserStatuses
        Active = 2
        Inactive = 1
        None = 0
    End Enum

#End Region

#Region " Connection Mode Initiation "

    Private Sub SetConnectionStatusColors(status As ConnectionStatuses)
        Dim foreColor As Color = Color.Black

        Select Case status

            Case ConnectionStatuses.Online
                Select Case Me.UserStatus
                    Case UserStatuses.Active
                        foreColor = Color.Green
                    Case UserStatuses.Inactive
                        foreColor = Color.FromArgb(100, 180, 100)
                    Case UserStatuses.None
                        foreColor = Color.DarkGray
                End Select

            Case ConnectionStatuses.Offline
                foreColor = Color.Red

            Case ConnectionStatuses.OnlineDegraded
                foreColor = Color.DarkSlateBlue

            Case ConnectionStatuses.OnlineSlow
                foreColor = Color.Magenta
        End Select

        Me.UIThread(Sub() Me.OwnerNode.ForeColor = foreColor)
    End Sub

    Private Sub SetRebootPendingFont()
        Select Case Me.ConnectionStatus

            Case ConnectionStatuses.Online
                If Not Me.RebootPending AndAlso Me.ConnectionStatus.Equals(ConnectionStatuses.Online) AndAlso Me.WMI IsNot Nothing Then
                    Try
                        Dim registryEntriesToCheckForRebootPending As New Dictionary(Of String, String) From
                                            {
                                                {"SYSTEM\CurrentControlSet\Control\Session Manager", "PENDINGFILERENAMEOPERATIONS"},
                                                {"SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update", "REBOOTREQUIRED"},
                                                {"Software\Microsoft\Windows\CurrentVersion\Component Based Servicing", "REBOOTPENDING"}
                                            }

                        Dim scope = If(Me.WMI.Architecture = ComputerArchitectures.X86, Me.WMI.X86Scope, Me.WMI.X64Scope)
                        Dim registry = New RegistryController(scope)

                        Dim registryKeys As String() = Nothing
                        For Each registryEntry As KeyValuePair(Of String, String) In registryEntriesToCheckForRebootPending
                            registryKeys = registry.GetKeyValues(registryEntry.Key, methodName:="EnumValues")
                            If registryKeys IsNot Nothing AndAlso registryKeys.Contains(registryEntry.Value) Then
                                Me.RebootPending = True
                                Exit Select
                            End If
                        Next

                    Catch ex As Exception
                        LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                    End Try
                End If

            Case Else
                Me.RebootPending = False

        End Select

        Me.UIThread(Sub() Me.OwnerNode.NodeFont = If(Me.RebootPending, NewFont(FontStyle.Italic), NewFont(FontStyle.Regular)))
    End Sub

    Private Function NewFont(style As FontStyle, Optional family As FontFamily = Nothing, Optional size As Single = 0) As Font
        If family Is Nothing Then
            family = Me.OwnerForm.ResourceExplorer.Font.FontFamily
        End If

        If size.Equals(0) Then
            size = Me.OwnerForm.ResourceExplorer.Font.Size
        End If

        Return New Font(family, size, style)
    End Function

    Private Sub DrawAreaDisposal()
        Dim garbage As New GarbageCollector()

        For Each control As Control In Me.SplitContainer.Panel1.Controls
            garbage.AddToGarbage = control
        Next

        Me.UIThread(Sub() garbage.DisposeAsync())
    End Sub

    Private Sub TabPage_Selecting(sender As Object, e As TabControlCancelEventArgs)
        If Not e.TabPage.Enabled Then
            e.Cancel = True
        End If
    End Sub

    Public Sub SetConnectionStatus(status As ConnectionStatuses)
        Me.ConnectionStatus = status
        Me.UserStatus = GetUserStatus()

        If Me.IsHandleCreated Then
            ' Clear the panel for Online Data
            Me.Initialized = True

            DrawAreaDisposal()
            SplitContainer.Panel1.InvokeClearControls()
            SetConnectionStatusColors(status)
            SetRebootPendingFont()

            Dim menuStrip As ContextMenuStrip = NewComputerMenuStrip(status)
            AddMenuStrip(menuStrip, Me.OwnerNode)

            Dim mainTabControl As New TabControl With
            {
                .Multiline = True,
                .Height = SplitContainer.Panel1.Height,
                .Width = SplitContainer.Panel1.Width,
                .Anchor = AnchorStyles.Bottom + AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Right
            }

            ' This event handler makes the tabs not "clickable" if disabled
            AddHandler mainTabControl.Selecting, AddressOf TabPage_Selecting

            ' Always add the main tab regardless of the connection status
            mainTabControl.TabPages.AddRange(New TabPage() {New MainTab(mainTabControl, Me, status)})

            ' Add all tabs for Online connection status
            If status = ConnectionStatuses.Online Then
                mainTabControl.TabPages.AddRange(New TabPage() _
                {
                    New AdvancedTab(mainTabControl, Me),
                    New ApplicationTab(mainTabControl, Me),
                    New UpdatesTab(mainTabControl, Me),
                    New PrinterTab(mainTabControl, Me),
                    New ProfileTab(mainTabControl, Me),
                    New ServicesTab(mainTabControl, Me),
                    New ProcessesTab(mainTabControl, Me)
                })
            End If

            ' Add the Tab Control to the Panel
            Me.SplitContainer.Panel1.InvokeAddControl(mainTabControl)

            WriteMessage(String.Format("Connection has completed loading (Status: {0})", [Enum].GetName(GetType(ConnectionStatuses), status)))

        Else
            Me.Initialized = False

            SetConnectionStatusColors(status)
            SetRebootPendingFont()
            Dim menuStrip As ContextMenuStrip = NewComputerMenuStrip(status)
            AddMenuStrip(menuStrip, OwnerNode)
        End If

    End Sub

    Public Enum ConnectionStatuses
        Offline = 0
        Online = 1
        OnlineDegraded = 2
        OnlineSlow = 4
    End Enum

#End Region

#Region " Remote Tool Initiation "

    Private Sub InitiateRemoteDesktop()
        Dim remoteDesktop As New RemoteTools(RemoteTools.RemoteTools.RemoteDesktop, Me)
        remoteDesktop.BeginWork()
    End Sub

    Private Sub InitiateRemoteAssistance()
        Dim remoteAssistance As New RemoteTools(RemoteTools.RemoteTools.RemoteAssistance, Me)
        remoteAssistance.BeginWork()
    End Sub

    Private Sub InitiateRemoteControlViewer()
        Dim remoteControlViewer As New RemoteTools(RemoteTools.RemoteTools.RemoteControlViewer, Me)
        remoteControlViewer.BeginWork()
    End Sub

    ''' <summary>
    ''' Remote Registry
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitiateRemoteRegistry()
        Dim remoteRegistry As New RemoteTools(RemoteTools.RemoteTools.RemoteRegistry, Me)
        remoteRegistry.BeginWork()
    End Sub

    Public Sub InitiateAdminShare(sender As Object, e As EventArgs)
        If Me.RespondsToPing() Then
            Dim adminShare As New RemoteTools(RemoteTools.RemoteTools.AdminShare, Me, sender.tag.ToString)
            adminShare.BeginWork()
        End If
    End Sub

    Public Sub InitiatePsExec()
        If Me.RespondsToPing() Then
            Dim psExec As New RemoteTools(RemoteTools.RemoteTools.PsExec, Me)
            psExec.BeginWork()
        End If
    End Sub

    Public Sub InitiateComputerManagement()
        If Me.RespondsToPing() Then
            Dim computerManagement As New RemoteTools(RemoteTools.RemoteTools.ComputerManagement, Me)
            computerManagement.BeginWork()
        End If
    End Sub

    Public Sub InitiateGroupPolicyEditor()
        If Me.RespondsToPing() Then
            Dim groupPolicyEditor As New RemoteTools(RemoteTools.RemoteTools.GroupPolicyEditor, Me)
            groupPolicyEditor.BeginWork()
        End If
    End Sub

    Public Sub InitiateResultantSetOfPolicy()
        If Me.RespondsToPing() Then
            Dim resultantSetOfPolicy As New RemoteTools(RemoteTools.RemoteTools.ResultantSetOfPolicy, Me)
            resultantSetOfPolicy.BeginWork()
        End If
    End Sub

    Public Sub InitiateCacExempt()
        If Me.RespondsToPing() Then
            Dim cacExempt As New RemoteTools(RemoteTools.RemoteTools.CACExempt, Me)
            cacExempt.BeginWork()
        End If
    End Sub

    Public Sub InitiateSetDescription()
        Dim setDescription As New RemoteTools(RemoteTools.RemoteTools.SetDescription, Me)
        setDescription.BeginWork()
    End Sub

    Public Sub InitiateSetLocation()
        Dim setLocation As New RemoteTools(RemoteTools.RemoteTools.SetLocation, Me)
        AddHandler setLocation.WorkCompleted, AddressOf LoadComputer
        setLocation.BeginWork()
    End Sub

    ''' <summary>
    ''' Toggle BitLocker
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitiateToggleBitLocker()
        If Me.RespondsToPing() Then
            Dim toggleBitLocker As New RemoteTools(RemoteTools.RemoteTools.ToggleBitLocker, Me)
            toggleBitLocker.BeginWork()
        End If
    End Sub

    ''' <summary>
    ''' Restart
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitiateRestart()
        If Me.RespondsToPing() Then
            Dim restart As New RemoteTools(RemoteTools.RemoteTools.Restart, Me)
            restart.BeginWork()
        End If
    End Sub

    ''' <summary>
    ''' No Prompt Restart
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitiateNoPromptRestart()
        If Me.RespondsToPing() Then
            Dim noPromptRestart As New RemoteTools(RemoteTools.RemoteTools.NoPromptRestart, Me)
            noPromptRestart.BeginWork()
        End If
    End Sub

    ''' <summary>
    ''' Logoff
    ''' </summary>
    ''' <remarks></remarks>
    Public Sub InitiateLogoff()
        If Me.RespondsToPing() Then
            Dim logoff As New RemoteTools(RemoteTools.RemoteTools.Logoff, Me)
            logoff.BeginWork()
        End If
    End Sub

    Public Sub InitiateCustomAction(sender As Object, e As EventArgs)
        Dim customAction As New RemoteTools(RemoteTools.RemoteTools.CustomAction, Me, sender.tag(0).ToString())
        customAction.BeginWork()
    End Sub

    Public Sub InitiateEnableComputers()
        Dim enableComputers As New RemoteTools(RemoteTools.RemoteTools.EnableComputers, Me)
        enableComputers.BeginWork()
    End Sub

    Public Sub InitiateDisableComputers()
        Dim disableComputers As New RemoteTools(RemoteTools.RemoteTools.DisableComputers, Me)
        disableComputers.BeginWork()
    End Sub

    Public Sub InitiateTimedRestart()
        Dim timedRestart As New RemoteTools(RemoteTools.RemoteTools.TimedRestart, Me)
        timedRestart.BeginWork()
    End Sub

#End Region

End Class