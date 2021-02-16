Imports System.ComponentModel

Public Class ServicesTab
    Inherits Tab

    Public Sub New(ownerTab As TabControl, computerPanel As ComputerPanel)
        MyBase.New(ownerTab, computerPanel)

        Me.Text = "Services"
        AddHandler InitWorker.DoWork, AddressOf InitializeServicesTab
        AddHandler ExportWorker.DoWork, AddressOf ExportServiceInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgDN As String = "Display Name"
        Public Shared ReadOnly Property lsvgSN As String = "Service Name"
        Public Shared ReadOnly Property lsvgSM As String = "Start Mode"
        Public Shared ReadOnly Property lsvgST As String = "State"
        Public Shared ReadOnly Property lsvgDE As String = "Description"
        Public Shared ReadOnly Property lsvgDS As String = "Dependent Services"
        Public Shared ReadOnly Property lsvgPN As String = "Process Name"
        Public Shared ReadOnly Property lsvgPI As String = "Process ID"
    End Structure

    Private Sub InitializeServicesTab()
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Dim serviceInfoListView As ListView = NewBaseListView(3)
        serviceInfoListView.Groups.Add(New ListViewGroup("lsvgSV", "Services"))

        For Each service As ManagementObject In ComputerPanel.WMI.Query("SELECT DisplayName, StartMode, State, Name FROM Win32_Service")
            If MyBase.UserCancellationPending() Then Exit Sub

            NewTabWriterItem(service.Properties("DisplayName").Value, New Object() {service.Properties("StartMode").Value, service.Properties(ListViewGroups.lsvgST).Value, service}, "lsvgSV")
        Next

        With serviceInfoListView
            .Items.AddRange(NewBaseListViewItems(serviceInfoListView, TabWriterObjects.ToArray()))
            .Sorting = SortOrder.Ascending
            .Sort()
        End With

        Dim serviceContainer As SplitContainer = NewSplitContainer(Me.Width - 250, "Choose a service from the list to see more information")
        Dim serviceSearchBox As TextBox = NewSearchTextBox("Enter your service name search here")

        AddHandler serviceInfoListView.SelectedIndexChanged, AddressOf FindServiceInfo

        Me.InvokeClearControls()
        Me.InvokeAddControl(serviceContainer)

        ShowListView(serviceInfoListView, serviceContainer, Panels.Panel1)
        AddHandler serviceInfoListView.ContextMenuStripChanged, AddressOf AddMenuStripOptions

        serviceSearchBox.Location = New Point(0, Me.Height - 25)
        Me.InvokeAddControl(serviceSearchBox)
    End Sub

    Private Sub FindServiceInfo()
        Me.LastSelectedListViewItem = GetSelectedListViewItem(Me.CurrentListView)

        If Not SelectionWorker.IsBusy Then
            SelectionWorker.RunWorkerAsync()
        End If
    End Sub

    Private Sub SelectionWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles SelectionWorker.DoWork
        Me.Controls(0).Controls(1).InvokeClearControls()

        MyBase.EnumWriterObjects = Nothing

        Dim selectedItem As ListViewItem = GetSelectedListViewItem(Me.CurrentListView)
        If selectedItem IsNot Nothing Then
            Dim serviceInfoListView As ListView = NewBaseListView(1)
            With serviceInfoListView.Groups
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgDN), ListViewGroups.lsvgDN))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgSN), ListViewGroups.lsvgSN))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgSM), ListViewGroups.lsvgSM))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgST), ListViewGroups.lsvgST))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgDE), ListViewGroups.lsvgDE))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgDS), ListViewGroups.lsvgDS))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPN), ListViewGroups.lsvgPN))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPI), ListViewGroups.lsvgPI))
            End With

            Dim selectedService As ManagementObject = selectedItem.Tag
            Dim queryText = $"SELECT DisplayName, Name, StartMode, State, Description, PathName, ProcessID FROM Win32_Service WHERE Name = '{selectedService.Properties("Name").Value}'"
            Dim services = Me.ComputerPanel.WMI.Query(queryText)
            If services IsNot Nothing Then
                For Each service As ManagementObject In services
                    MyBase.NewEnumWriterItem(service.Properties("DisplayName").Value, service, NameOf(ListViewGroups.lsvgDN))
                    MyBase.NewEnumWriterItem(service.Properties("Name").Value, service, NameOf(ListViewGroups.lsvgSN))
                    MyBase.NewEnumWriterItem(service.Properties("StartMode").Value, service, NameOf(ListViewGroups.lsvgSM))
                    MyBase.NewEnumWriterItem(service.Properties("State").Value, service, NameOf(ListViewGroups.lsvgST))

                    Dim description As String = service.Properties("Description").Value
                    If Not String.IsNullOrWhiteSpace(description) Then
                        MyBase.NewEnumWriterItem(description, service, NameOf(ListViewGroups.lsvgDE))
                    End If

                    Dim dependentServices As String() = New ServiceController(Me.ComputerPanel.WMI).CheckDependentServices(service.Properties("Name").Value)
                    If dependentServices IsNot Nothing AndAlso dependentServices.Count > 0 Then
                        For Each dependentService As String In dependentServices
                            If Not String.IsNullOrWhiteSpace(dependentService) Then
                                Dim managementObjects = Me.ComputerPanel.WMI.Query($"SELECT DisplayName FROM Win32_Service WHERE Name = '{dependentService}'")
                                MyBase.NewEnumWriterItem(Me.ComputerPanel.WMI.GetPropertyValue(managementObjects, "DisplayName"), service, NameOf(ListViewGroups.lsvgDS))
                            End If
                        Next
                    End If
                    MyBase.NewEnumWriterItem(service.Properties("PathName").Value, service, NameOf(ListViewGroups.lsvgPN))

                    If New ServiceController(Me.ComputerPanel.WMI).QueryState(service.Properties("Name").Value) = ServiceController.ServiceState.Running Then
                        MyBase.NewEnumWriterItem(service.Properties("ProcessID").Value, service, NameOf(ListViewGroups.lsvgPI))
                    End If
                Next
            End If

            serviceInfoListView.Items.AddRange(NewBaseListViewItems(serviceInfoListView, MyBase.EnumWriterObjects.ToArray()))

            Dim serviceSplitContainer As New SplitContainer()
            serviceSplitContainer.Orientation = Orientation.Horizontal

            If GetSelectedListViewItem(Me.CurrentListView) Is selectedItem Then
                ShowListView(serviceInfoListView, Me.Controls(0), Panels.Panel2)
            Else
                SelectionWorker_DoWork(sender, e)
            End If
        End If
    End Sub

    Private Sub AddMenuStripOptions()
        If MyBase.CurrentListView.SelectedItems.Count > 0 Then
            Dim service As ManagementObject = GetSelectedListViewItem(MyBase.CurrentListView).Tag
            Dim serviceController As New ServiceController(Me.ComputerPanel.WMI)
            Dim setStartupTypeMenuItem As New ToolStripMenuItem("Set Startup Type")

            Select Case serviceController.QueryStartupType(service.Properties("Name").Value)
                Case ServiceController.ServiceStartupType.Auto
                    setStartupTypeMenuItem.DropDownItems.Add("Manual", Nothing, AddressOf SetServiceManual)
                    setStartupTypeMenuItem.DropDownItems.Add("Disabled", Nothing, AddressOf SetServiceDisable)
                Case ServiceController.ServiceStartupType.Manual
                    setStartupTypeMenuItem.DropDownItems.Add("Auto", Nothing, AddressOf SetServiceAuto)
                    setStartupTypeMenuItem.DropDownItems.Add("Disabled", Nothing, AddressOf SetServiceDisable)
                Case ServiceController.ServiceStartupType.Disabled
                    setStartupTypeMenuItem.DropDownItems.Add("Auto", Nothing, AddressOf SetServiceAuto)
                    setStartupTypeMenuItem.DropDownItems.Add("Manual", Nothing, AddressOf SetServiceManual)
            End Select

            AddToCurrentMenuStrip(setStartupTypeMenuItem)
            AddToCurrentMenuStrip(New ToolStripSeparator)

            Select Case serviceController.QueryState(service.Properties("Name").Value)
                Case ServiceController.ServiceState.Running
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Stop Service", Nothing, AddressOf StopService))
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Restart Service", Nothing, AddressOf RestartService))
                Case ServiceController.ServiceState.Stopped
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Start Service", Nothing, AddressOf StartService))
            End Select

            AddToCurrentMenuStrip(New ToolStripSeparator)
        End If
    End Sub

    Private Sub SetServiceAuto()
        Dim serviceSetAuto As New RemoteTools(RemoteTools.RemoteTools.ServiceSetAuto, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(Me.CurrentListView).Tag})
        AddHandler serviceSetAuto.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        serviceSetAuto.BeginWork()
    End Sub

    Private Sub SetServiceManual()
        Dim serviceSetManual As New RemoteTools(RemoteTools.RemoteTools.ServiceSetManual, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(Me.CurrentListView).Tag})
        AddHandler serviceSetManual.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        serviceSetManual.BeginWork()
    End Sub

    Private Sub SetServiceDisable()
        Dim serviceSetDisable As New RemoteTools(RemoteTools.RemoteTools.ServiceSetDisable, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(Me.CurrentListView).Tag})
        AddHandler serviceSetDisable.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        serviceSetDisable.BeginWork()
    End Sub

    Private Sub StopService()
        Dim serviceStop As New RemoteTools(RemoteTools.RemoteTools.ServiceStop, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(Me.CurrentListView).Tag})
        AddHandler serviceStop.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        serviceStop.BeginWork()
    End Sub

    Private Sub StartService()
        Dim serviceStart As New RemoteTools(RemoteTools.RemoteTools.ServiceStart, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(Me.CurrentListView).Tag})
        AddHandler serviceStart.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        serviceStart.BeginWork()
    End Sub

    Private Sub RestartService()
        Dim serviceRestart As New RemoteTools(RemoteTools.RemoteTools.ServiceRestart, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(Me.CurrentListView).Tag})
        AddHandler serviceRestart.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        serviceRestart.BeginWork()
    End Sub

    Private Sub ExportServiceInfo(sender As Object, e As DoWorkEventArgs)
        ' Query the host's WMI provider for a collection of Win32_Service objects
        Dim queryText = "SELECT DisplayName, Name, StartMode, State, Description, PathName, ProcessID FROM Win32_Service"
        Dim services = Me.ComputerPanel.WMI.Query(queryText).Cast(Of ManagementObject)()

        ' Get the total count of objects for calcuating progress
        Dim servicesCount As Integer = services.Count

        ' Set the maximum value for the progress bar
        Me.UIThread(Sub() MyBase.ExportingProgressBar.Maximum = servicesCount)

        ' Create a list of ProcessInfo objects from the raw tab data
        Dim serviceInfos As New List(Of ServiceInfo)
        For index As Integer = 0 To servicesCount - 1
            If MyBase.ExportWorker.CancellationPending Then Exit For

            Dim state As String = services(index).Properties("State").Value
            Dim processID As String = services(index).Properties("ProcessID").Value

            serviceInfos.Add(New ServiceInfo() With
            {
                .DisplayName = services(index).Properties("DisplayName").Value,
                .ServiceName = services(index).Properties("Name").Value,
                .StartMode = services(index).Properties("StartMode").Value,
                .State = state,
                .Description = services(index).Properties("Description").Value,
                .ProcessName = services(index).Properties("PathName").Value,
                .ProcessID = If(state = "Running", processID, String.Empty)
            })

            MyBase.ExportWorker.ReportProgress(100 * (index / servicesCount))
        Next

        ' Write the list out to the CSV file chosen by the user in the base class Tab.BeginExportProcess()
        If serviceInfos.Count > 0 Then
            Dim userSelectedCSVFilePath As String = e.Argument
            WriteListOfObjectsToCSV(serviceInfos.OrderBy(Function(srvInfo) srvInfo.DisplayName), userSelectedCSVFilePath, includeHeader:=True)
        End If
    End Sub

End Class
