Imports System.Reflection
Imports System.ComponentModel

Public Class ApplicationTab
    Inherits Tab

    Private Property EnumWriterApps As List(Of Object)
    Private WithEvents AppBackgroundThread As New BackgroundWorker()

    Public Sub New(ownerTab As TabControl, computerContext As ComputerControl)
        MyBase.New(ownerTab, computerContext)

        Me.Text = "Applications"
        AddHandler MyBase.LoaderBackgroundThread.DoWork, AddressOf Me.InitializeAppsTab
        AddHandler MyBase.ExportBackgroundThread.DoWork, AddressOf Me.ExportAppInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgAN As String = "Applications"
        Public Shared ReadOnly Property lsvgSN As String = "Software Name"
        Public Shared ReadOnly Property lsvgRK As String = "Registry Key"
        Public Shared ReadOnly Property lsvgRI As String = "Registry Information"
        Public Shared ReadOnly Property lsvgVI As String = "Version Information"
        Public Shared ReadOnly Property lsvgUS As String = "Uninstall String"
        Public Shared ReadOnly Property lsvgID As String = "Install Date"
        Public Shared ReadOnly Property lsvgIS As String = "Install Source"
    End Structure

    Private Sub InitializeAppsTab()
        ' Clear the tab of all child controls
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Dim appInfoListView As New ListView()
        appInfoListView = NewBasicInfoListView(1)
        appInfoListView.Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgAN), ListViewGroups.lsvgAN))

        Dim x86Registry As New RegistryController(Me.ComputerContext.WMI.X86Scope)
        For Each keyValue In x86Registry.GetKeyValues("Software\Microsoft\Windows\CurrentVersion\Uninstall")
            If MyBase.UserCancellationPending() Then Exit Sub

            Dim registryKeyPath As String = String.Format("Software\Microsoft\Windows\CurrentVersion\Uninstall\{0}", keyValue)
            Dim displayName As String = x86Registry.GetKeyValue(registryKeyPath, "DisplayName", RegistryController.RegistryKeyValueTypes.String)

            If String.IsNullOrWhiteSpace(displayName) Then
                Continue For
            End If

            NewTabWriterItem(displayName, String.Format("Software\Microsoft\Windows\CurrentVersion\Uninstall\{0}&32", keyValue), NameOf(ListViewGroups.lsvgAN))
        Next

        If Me.ComputerContext.WMI.X64Scope.IsConnected Then
            Dim x64Registry As New RegistryController(Me.ComputerContext.WMI.X64Scope)
            For Each keyValue In x64Registry.GetKeyValues("Software\Microsoft\Windows\CurrentVersion\Uninstall")
                If MyBase.UserCancellationPending() Then Exit Sub

                Dim registryKeyPath As String = String.Format("Software\Microsoft\Windows\CurrentVersion\Uninstall\{0}", keyValue)
                Dim displayName As String = x64Registry.GetKeyValue(registryKeyPath, "DisplayName", RegistryController.RegistryKeyValueTypes.String)

                If String.IsNullOrWhiteSpace(displayName) Then
                    Continue For
                End If

                NewTabWriterItem(displayName, String.Format("Software\Microsoft\Windows\CurrentVersion\Uninstall\{0}&64", keyValue), NameOf(ListViewGroups.lsvgAN))
            Next
        End If

        appInfoListView.Items.AddRange(NewBaseListViewItems(appInfoListView, TabWriterObjects.ToArray()))
        appInfoListView.Sorting = SortOrder.Ascending
        appInfoListView.Sort()

        ' Create the Split Container that will house the Application Listing and the Application information
        Dim applicationSplitContainer As SplitContainer = NewSplitContainer(Me.Width - 350, "Choose an application from the list provided to see more information")
        Dim searchTextBox As TextBox = NewSearchTextBox("Enter your application search here")

        ' Clear all controls from the tab and add the split container to the tab
        Me.InvokeClearControls()
        Me.InvokeAddControl(applicationSplitContainer)

        ' Create a handler for selection changes in the Application list that will allow information to be generated.
        AddHandler appInfoListView.SelectedIndexChanged, AddressOf GetAppInformation

        ' Display the Application listing in the first panel of the split container
        ShowListView(appInfoListView, applicationSplitContainer, Panels.Panel1)

        ' Place the search textbox in the bottom of the tab
        searchTextBox.Location = New Point(0, Me.Height - 25)

        Me.InvokeAddControl(searchTextBox)
        Me.UIThread(Sub() Me.MainListView.Focus())
    End Sub

    Public Sub NewAppWriteItem(subject As String, body As Object, group As String)
        If Me.EnumWriterApps Is Nothing Then
            Me.EnumWriterApps = New List(Of Object)()
        End If

        Me.EnumWriterApps.Add(New Object() {String.Format("   {0}", subject), group, body})
    End Sub

    Private Sub GetAppInformation()
        If Not Me.AppBackgroundThread.IsBusy Then
            Me.AppBackgroundThread.RunWorkerAsync()
        End If
    End Sub

    Private Sub AppBackgroundThread_DoWork(sender As Object, e As DoWorkEventArgs) Handles AppBackgroundThread.DoWork
        If Me.ComputerContext.WMI.X86Scope.IsConnected Then
            Me.Controls(0).Controls(1).InvokeClearControls()

            Me.EnumWriterApps = Nothing

            Dim applicationListView = Me.MainListView
            Dim selectedItem As ListViewItem = GetSelectedListViewItem(applicationListView)

            Dim appInfoListView As ListView = NewBasicInfoListView(1)
            With appInfoListView.Groups
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgSN), ListViewGroups.lsvgSN))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgRK), ListViewGroups.lsvgRK))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgRI), ListViewGroups.lsvgRI))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgVI), ListViewGroups.lsvgVI))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgUS), ListViewGroups.lsvgUS))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgID), ListViewGroups.lsvgID))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgIS), ListViewGroups.lsvgIS))
            End With

            Try
                Dim rawKeys As String() = selectedItem.SubItems(1).Text.Split("&")
                Dim registry As RegistryController = Nothing

                If rawKeys(1) = "32" Then registry = New RegistryController(Me.ComputerContext.WMI.X86Scope)
                If rawKeys(1) = "64" Then registry = New RegistryController(Me.ComputerContext.WMI.X64Scope)

                NewAppWriteItem(registry.GetKeyValue(rawKeys(0), "DisplayName", RegistryController.RegistryKeyValueTypes.String), " ", NameOf(ListViewGroups.lsvgSN))
                NewAppWriteItem(String.Format("HKLM\{0}", rawKeys(0)), " ", NameOf(ListViewGroups.lsvgRK))
                NewAppWriteItem(String.Format("{0}-bit Registry", rawKeys(1)), " ", NameOf(ListViewGroups.lsvgRI))

                Dim displayVersion As String = registry.GetKeyValue(rawKeys(0), "DisplayVersion", RegistryController.RegistryKeyValueTypes.String)
                If Not String.IsNullOrWhiteSpace(displayVersion) Then
                    NewAppWriteItem(displayVersion, " ", NameOf(ListViewGroups.lsvgVI))
                End If

                Dim uninstallString As String = registry.GetKeyValue(rawKeys(0), "UninstallString", RegistryController.RegistryKeyValueTypes.String)
                If Not String.IsNullOrWhiteSpace(uninstallString) Then
                    NewAppWriteItem(uninstallString, " ", NameOf(ListViewGroups.lsvgUS))
                End If

                Dim installDate As String = registry.GetKeyValue(rawKeys(0), "InstallDate", RegistryController.RegistryKeyValueTypes.String)
                If Not String.IsNullOrWhiteSpace(installDate) Then NewAppWriteItem(installDate, " ", NameOf(ListViewGroups.lsvgID))

                Dim installSource As String = registry.GetKeyValue(rawKeys(0), "InstallSource", RegistryController.RegistryKeyValueTypes.String)
                If Not String.IsNullOrWhiteSpace(installSource) Then
                    NewAppWriteItem(installSource, " ", NameOf(ListViewGroups.lsvgIS))
                End If

                appInfoListView.Items.AddRange(NewBaseListViewItems(appInfoListView, Me.EnumWriterApps.ToArray()))
                Me.Controls(0).Controls(1).InvokeClearControls()

                If GetSelectedListViewItem(applicationListView) Is selectedItem Then
                    ShowListView(appInfoListView, Me.Controls(0), Panels.Panel2)
                Else
                    AppBackgroundThread_DoWork(sender, e)
                End If

            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))

                If Not Me.ComputerContext.RespondsToPing() Then
                    Me.ComputerContext.SetConnectionStatus(ComputerControl.ConnectionStatuses.Offline)
                Else
                    ValidateWMI()
                End If
            End Try
        End If
    End Sub

    Private Sub ExportAppInfo(sender As Object, e As DoWorkEventArgs)
        ' Get the total count of objects for calcuating progress
        Dim tabWriterObjectsCount As Integer = MyBase.TabWriterObjects.Count()

        ' Set the maximum value for the progress bar
        Me.UIThread(Sub() MyBase.ExportingProgressBar.Maximum = tabWriterObjectsCount)

        ' Create a list of ApplicationInfo objects from the raw tab data
        Dim applicationInfos As New List(Of ApplicationInfo)
        For index As Integer = 0 To tabWriterObjectsCount - 1
            If MyBase.ExportBackgroundThread.CancellationPending Then Exit For

            Dim appKeyPathAndArchitecture As String() = MyBase.TabWriterObjects(index)(2).ToString().Split("&")
            Dim appKeyPath As String = appKeyPathAndArchitecture(0)
            Dim appArchitecture As String = appKeyPathAndArchitecture(1)

            Dim registry As New RegistryController()
            If appArchitecture = "32" Then registry = New RegistryController(ComputerContext.WMI.X86Scope)
            If appArchitecture = "64" Then registry = New RegistryController(ComputerContext.WMI.X64Scope)

            applicationInfos.Add(New ApplicationInfo() With
                {
                    .SoftwareName = registry.GetKeyValue(appKeyPath, "DisplayName", RegistryController.RegistryKeyValueTypes.String),
                    .RegistryKey = appKeyPath,
                    .RegistryInformation = String.Format("{0}-bit Registry", appArchitecture),
                    .Version = registry.GetKeyValue(appKeyPath, "DisplayVersion", RegistryController.RegistryKeyValueTypes.String),
                    .UninstallString = registry.GetKeyValue(appKeyPath, "UninstallString", RegistryController.RegistryKeyValueTypes.String),
                    .InstallDate = registry.GetKeyValue(appKeyPath, "InstallDate", RegistryController.RegistryKeyValueTypes.String),
                    .InstallSource = registry.GetKeyValue(appKeyPath, "InstallSource", RegistryController.RegistryKeyValueTypes.String)
                })

            MyBase.ExportBackgroundThread.ReportProgress(100 * (index / tabWriterObjectsCount))
        Next

        ' Write the list out to the CSV file chosen by the user in the base class Tab.BeginExportProcess()
        If applicationInfos.Count > 0 Then
            Dim userSelectedCSVFilePath As String = e.Argument
            WriteListOfObjectsToCSV(applicationInfos.OrderBy(Function(appInfo) appInfo.SoftwareName), userSelectedCSVFilePath, includeHeader:=True)
        End If
    End Sub

End Class