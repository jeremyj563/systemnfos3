Imports System.ComponentModel

Public Class ProcessesTab
    Inherits Tab

    Public Sub New(ownerTab As TabControl, computerPanel As ComputerPanel)
        MyBase.New(ownerTab, computerPanel)

        Me.Text = "Processes"
        AddHandler InitWorker.DoWork, AddressOf InitializeProcessesTab
        AddHandler ExportWorker.DoWork, AddressOf ExportProcessInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgPC As String = "Processes"
        Public Shared ReadOnly Property lsvgPN As String = "Process Name"
        Public Shared ReadOnly Property lsvgPI As String = "Process ID"
        Public Shared ReadOnly Property lsvgCD As String = "Creation Date"
        Public Shared ReadOnly Property lsvgTC As String = "Thread Count"
        Public Shared ReadOnly Property lsvgCL As String = "Command Line"
        Public Shared ReadOnly Property lsvgNI As String = "No Information Available"
    End Structure

    Private Sub InitializeProcessesTab()
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Dim processInfoListView As New ListView()
        processInfoListView = NewBasicInfoListView(2)
        processInfoListView.Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPC), ListViewGroups.lsvgPC))

        For Each process As ManagementObject In Me.ComputerPanel.WMI.Query("SELECT Name,ProcessID,CreationDate FROM Win32_Process")
            If MyBase.UserCancellationPending() Then Exit Sub

            NewTabWriterItem(process.Properties("Name").Value, New Object() {process.Properties("ProcessID").Value, process}, NameOf(ListViewGroups.lsvgPC))
        Next

        Dim processesSplitContainer As SplitContainer = NewSplitContainer(Me.Width - 450, "Choose a process from the list to see more information")

        With processInfoListView
            .Items.AddRange(NewBaseListViewItems(MainListView, TabWriterObjects.ToArray()))
            .Sorting = SortOrder.Ascending
            .Sort()
        End With

        Dim processSearchTextBox As TextBox = NewSearchTextBox("Enter your process name search here")
        Me.InvokeClearControls()
        Me.InvokeAddControl(processesSplitContainer)
        ShowListView(Me.MainListView, processesSplitContainer, Panels.Panel1)

        AddHandler processInfoListView.ContextMenuStripChanged, AddressOf AddMenuStripOptions

        processSearchTextBox.Location = New Point(0, Me.Height - 25)
        Me.InvokeAddControl(processSearchTextBox)

        AddHandler processInfoListView.SelectedIndexChanged, AddressOf FindProcessInformation
    End Sub

    Private Sub FindProcessInformation()
        Me.LastSelectedListViewItem = GetSelectedListViewItem(Me.MainListView)

        If Not MyBase.SelectionWorker.IsBusy Then
            MyBase.SelectionWorker.RunWorkerAsync()
        End If
    End Sub

    Private Sub SelectionWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles SelectionWorker.DoWork
        Me.Controls(0).Controls(1).InvokeClearControls()

        MyBase.TabWriterObjects = Nothing

        Dim selectedItem As ListViewItem = GetSelectedListViewItem(Me.MainListView)

        If selectedItem IsNot Nothing Then
            Dim process As ManagementObject = selectedItem.Tag
            Dim processInfoListView As ListView = NewBasicInfoListView(1)
            With processInfoListView.Groups
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPN), ListViewGroups.lsvgPN))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPI), ListViewGroups.lsvgPI))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgCD), ListViewGroups.lsvgCD))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgTC), ListViewGroups.lsvgTC))
                .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgCL), ListViewGroups.lsvgCL))
            End With

            Dim queryText = $"SELECT Name, ProcessID, CreationDate, ThreadCount, CommandLine FROM Win32_Process WHERE ProcessID = '{process.Properties("ProcessID").Value}'"
            For Each proc In Me.ComputerPanel.WMI.Query(queryText)
                MyBase.NewTabWriterItem(proc.Properties("Name").Value, proc, NameOf(ListViewGroups.lsvgPN))
                MyBase.NewTabWriterItem(proc.Properties("ProcessID").Value, proc, NameOf(ListViewGroups.lsvgPI))
                MyBase.NewTabWriterItem(proc.Properties("CreationDate").Value, proc, NameOf(ListViewGroups.lsvgCD))
                MyBase.NewTabWriterItem(proc.Properties("ThreadCount").Value, proc, NameOf(ListViewGroups.lsvgTC))
                MyBase.NewTabWriterItem(proc.Properties("CommandLine").Value, proc, NameOf(ListViewGroups.lsvgCL))

                If New ProcessController(Me.ComputerPanel.WMI).QueryProcessState(proc.Properties("Name").Value) = ProcessController.ProcessCondition.SingleRunning Then
                    NewEnumWriterItem(proc.Properties("ProcessID").Value, proc, NameOf(ListViewGroups.lsvgPN))
                End If
            Next

            If MyBase.TabWriterObjects IsNot Nothing Then
                processInfoListView.Items.AddRange(NewBaseListViewItems(processInfoListView, MyBase.TabWriterObjects.ToArray()))
            Else
                processInfoListView.Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgNI), ListViewGroups.lsvgNI))
                MyBase.NewTabWriterItem(" ", New ManagementObject, NameOf(ListViewGroups.lsvgNI))
            End If

            If GetSelectedListViewItem(Me.MainListView) Is selectedItem Then
                ShowListView(processInfoListView, Me.Controls(0), Panels.Panel2)
            Else
                SelectionWorker_DoWork(sender, e)
            End If

        End If
    End Sub

    Private Sub AddMenuStripOptions()
        If Me.MainListView.SelectedItems.Count > 0 Then
            Dim selectedItem As ManagementObject = GetSelectedListViewItem(MainListView).Tag
            Dim process As New ProcessController(Me.ComputerPanel.WMI)

            Select Case process.QueryProcessState(selectedItem.Properties("Name").Value)
                Case ProcessController.ProcessCondition.SingleRunning
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Stop Process", Nothing, AddressOf StopProcess))
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                Case ProcessController.ProcessCondition.MultipleRunning
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Stop Process", Nothing, AddressOf StopProcess))
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Stop All Processes", Nothing, AddressOf StopAllProcesses))
                    AddToCurrentMenuStrip(New ToolStripSeparator)
            End Select
        End If
    End Sub

    Private Sub StopProcess()
        Dim processStop As New RemoteTools(RemoteTools.RemoteTools.ProcessStop, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(MainListView).Tag})
        AddHandler processStop.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        processStop.BeginWork()
    End Sub

    Private Sub StopAllProcesses()
        Dim processes = New ProcessController(Me.ComputerPanel.WMI).GetProcesses(GetSelectedListViewItem(MainListView).Tag.Properties("Name").Value)
        For Each process In processes
            Dim processStop As New RemoteTools(RemoteTools.RemoteTools.ProcessStop, Me.ComputerPanel, {process})
            AddHandler processStop.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
            processStop.BeginWork()
        Next
    End Sub

    Private Sub ExportProcessInfo(sender As Object, e As DoWorkEventArgs)
        ' Query the host's WMI provider for a collection of Win32_Process objects
        Dim queryText = "SELECT Name, ProcessID, CreationDate, ThreadCount, CommandLine FROM Win32_Process"
        Dim processes = Me.ComputerPanel.WMI.Query(queryText).Cast(Of ManagementObject)()

        ' Get the total count of objects for calcuating progress
        Dim processesCount As Integer = processes.Count

        ' Set the maximum value for the progress bar
        Me.UIThread(Sub() MyBase.ExportingProgressBar.Maximum = processesCount)

        ' Create a list of ProcessInfo objects from the raw tab data
        Dim processInfos As New List(Of ProcessInfo)
        For index As Integer = 0 To processesCount - 1
            If MyBase.ExportWorker.CancellationPending Then Exit For

            processInfos.Add(New ProcessInfo With
            {
                .ProcessName = processes(index).Properties("Name").Value,
                .ProcessID = processes(index).Properties("ProcessID").Value,
                .CreationDate = processes(index).Properties("CreationDate").Value,
                .ThreadCount = processes(index).Properties("ThreadCount").Value,
                .CommandLine = processes(index).Properties("CommandLine").Value
            })

            MyBase.ExportWorker.ReportProgress(100 * (index / processesCount))
        Next

        ' Write the list out to the CSV file chosen by the user in the base class Tab.BeginExportProcess()
        If processInfos.Count > 0 Then
            Dim userSelectedCSVFilePath As String = e.Argument
            WriteListOfObjectsToCSV(processInfos.OrderBy(Function(procInfo) procInfo.ProcessName), userSelectedCSVFilePath, includeHeader:=True)
        End If
    End Sub

End Class