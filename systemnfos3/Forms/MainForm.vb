Imports System.ComponentModel
Imports System.Reflection
Imports System.Threading
Imports System.Text
Imports System.IO
Imports Microsoft.Win32
Imports System.Threading.Tasks
Imports Newtonsoft.Json
Imports System.Net
Imports System.Web

Public Class MainForm

    Public Property BindingSource As New BindingSource()
    Private ReadOnly Property DefaultTitle As String = GetDefaultTitle()
    Private Property MouseDownOnListView As Boolean = False
    Private Property LastChangedLDAPTime As Date
    Private Property ImportWorker As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
    Private Property LDAPWorker As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
    Private Property AppUpdates As NuGetSearchResult
    Private Property AppUpdateNotificationSent As Boolean

    Private ldapLock As New Object
    Private statusStripLock As New Object
    Private updateLock As New Object
    Private nodeLock As New Object
    Private settingsLock As New Object

#Region " Initialization "

    Public Sub New(bindingList As BindingList(Of DataUnit), lastChangedLDAPTime As Date)
        ' This call is required by the designer
        InitializeComponent()

        ' Configure the BindingSource
        With Me.BindingSource
            .DataSource = bindingList
            .DataSource.AllowEdit = True
            .DataSource.AllowNew = True
            .DataSource.AllowRemove = True
        End With

        ' Set the timestamp for the last changed container in ldap.
        ' This will be used as a launching timestamp for incremental ldap updates.
        Me.LastChangedLDAPTime = lastChangedLDAPTime

        ' Configure the user input box
        With Me.UserInputComboBox
            .ValueMember = NameOf(DataUnit.Value)
            .DisplayMember = NameOf(DataUnit.Display)
            .DataSource = Me.BindingSource
            .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            .AutoCompleteSource = AutoCompleteSource.ListItems
            .SelectedItem = Nothing
            .Focus()
        End With

        ' Get the root tree node from the resource explorer
        Dim rootNode As TreeNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode))

        ' Populate SysTool collections in the TreeView
        Dim collectionNodeMenuStrip As New ContextMenuStrip()
        collectionNodeMenuStrip.Items.Add("Delete", Nothing, AddressOf DeleteCollection)

        For index As Integer = (bindingList.Count - 1) To 0 Step -1
            If TypeOf bindingList(index) Is Collection Then
                Dim treeNode As New TreeNode() With
                    {
                        .Name = bindingList(index).Value,
                        .Text = bindingList(index).Value,
                        .Tag = bindingList(index).Value,
                        .ContextMenuStrip = collectionNodeMenuStrip
                    }

                rootNode.Nodes(NameOf(Nodes.Collections)).Nodes.Add(treeNode)
            ElseIf TypeOf bindingList(index) Is Computer Then
                Exit For
            End If
        Next

        ' Load Custom Actions from Registry into the TreeView
        Dim registry As New RegistryController()

        For Each customAction In registry.GetKeyValues(My.Settings.RegistryPathCustomActions, RegistryHive.CurrentUser)
            Dim treeNode As New TreeNode()
            Dim customActionMenuStrip As New ContextMenuStrip()
            With treeNode
                .Name = customAction
                .Text = customAction
                .Tag = New CustomActionsControl(Me, customAction, treeNode)
                .ContextMenuStrip = customActionMenuStrip
            End With

            Me.ResourceExplorer.Nodes(NameOf(Nodes.Settings)).Nodes(NameOf(Nodes.CustomActions)).Nodes.Add(treeNode)
            customActionMenuStrip.Items.Add("Delete", Nothing, Sub() DeleteCustomAction(treeNode, customAction))
        Next

        ' Custom Actions options menu
        Dim customActionsMenuStrip As New ContextMenuStrip()
        customActionsMenuStrip.Items.Add("New", Nothing, AddressOf NewCustomAction)
        Me.ResourceExplorer.Nodes(NameOf(Nodes.Settings)).Nodes(NameOf(Nodes.CustomActions)).ContextMenuStrip = customActionsMenuStrip

        ' Resource Explorer options menu
        Dim resourceExplorerMenuStrip As New ContextMenuStrip()
        resourceExplorerMenuStrip.Items.Add("Change Main Window Title", Nothing, AddressOf SetMainWindowTitle)
        rootNode.ContextMenuStrip = resourceExplorerMenuStrip

        ' Collections options menu
        Dim collectionsMenuStrip As New ContextMenuStrip()
        collectionsMenuStrip.Items.Add("New Collection", Nothing, AddressOf NewCollection)
        rootNode.Nodes(NameOf(Nodes.Collections)).ContextMenuStrip = collectionsMenuStrip

        ' After loading all menus, expand the first menu (Resource Explorer)
        rootNode.Expand()

        ' Set the submit button as the global accept button if the ComboBox is being used
        AddHandler Me.UserInputComboBox.TextChanged, Sub() SetAcceptButton(Me.SubmitButton)
        AddHandler Me.UserInputComboBox.Click, Sub() SetAcceptButton(Me.SubmitButton)

        ' Add event handler to begin background ping monitoring of loaded computer nodes
        AddHandler Me.Shown, AddressOf RunPingMonitor

        ' Add event handler to begin background up time monitoring
        AddHandler Me.Shown, AddressOf RunStatusStripMonitor

        ' Add event handler to begin monitoring for program version updates
        AddHandler Me.Shown, AddressOf RunUpdateMonitor

        ' Add event handler to begin background ldap incremental updates
        'AddHandler Me.LDAPWorker.DoWork, AddressOf UpdateLDAPDataBindings
        'AddHandler Me.LDAPWorker.RunWorkerCompleted, Sub(s, e) If Not e.Cancelled Then Me.LDAPWorker.RunWorkerAsync()
        'AddHandler Me.FormClosing, AddressOf LDAPWorker.CancelAsync

        ' Add event handlers for the import functionality
        AddHandler Me.ImportWorker.DoWork, AddressOf MassImportProcessComputers
        AddHandler Me.ImportWorker.RunWorkerCompleted, AddressOf RefreshComputerListView
        AddHandler Me.FormClosing, AddressOf ImportWorker.CancelAsync

        ' Create the main ContextMenuStrip for the Computer TreeNode
        Dim computersMenuStrip As New ContextMenuStrip()
        With computersMenuStrip.Items
            .Add("Clear All", Nothing, AddressOf RemoveAllComputerNodes)
            .Add(New ToolStripSeparator())
        End With

        Dim importMenuItem As New ToolStripMenuItem("Import Computers")
        With importMenuItem.DropDownItems
            .Add("From Clipboard", Nothing, AddressOf MassImportFromClipboard)
            .Add("From Text File", Nothing, AddressOf MassImportFromTextFile)
        End With

        computersMenuStrip.Items.Add(importMenuItem)
        AddHandler computersMenuStrip.Opening, AddressOf PopulateDynamicMenuItems
        AddHandler computersMenuStrip.Closing, AddressOf RemoveDynamicMenuItems
        rootNode.Nodes(NameOf(Nodes.Computers)).ContextMenuStrip = computersMenuStrip

        ' Load any computers that are saved in user settings
        UserSettingsLoadComputers()
    End Sub

    Private Function GetDefaultTitle() As String
        Dim title As String = Nothing

        Try
            title = $"System Tool 3 - {Environment.UserName}@{Environment.MachineName} - !!! TEST ATTEMPT 9 !!!"
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return title
    End Function

#End Region

#Region " Main Form "

    Private Sub Me_Load(sender As Object, e As EventArgs) Handles Me.Load
        [Global].MainForm = Me
        [Global].SetControlDoubleBufferedProperty(Me.ResourceExplorer)
        UserSettingsLoadFormSize()
    End Sub

    Private Sub Me_Shown(sender As Object, e As EventArgs) Handles Me.Shown
        Me.UserInputComboBox.Text = String.Empty
        Me.LDAPWorker.RunWorkerAsync()
    End Sub

    Private Sub Me_ResizeEnd(sender As Object, e As EventArgs) Handles Me.ResizeEnd
        UserSettingsSaveFormSize()
    End Sub

    Private Sub ClearButton_Click(sender As Object, e As EventArgs) Handles ClearButton.Click
        Me.ResourceExplorer.SelectedNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode))
        With Me.UserInputComboBox
            .Text = String.Empty
            .Focus()
        End With
    End Sub

    Private Sub SubmitButton_Click(sender As Object, e As EventArgs) Handles SubmitButton.Click
        ' Determines the type of class contained within the Combobox selection and handles data accordingly.
        Me.AcceptButton = Nothing
        If Me.UserInputComboBox.SelectedItem IsNot Nothing Then
            Select Case True
                Case TypeOf Me.UserInputComboBox.SelectedItem Is Collection
                    Dim collection As Collection = UserInputComboBox.SelectedItem
                    Me.ResourceExplorer.SelectedNode = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Collections)).Nodes(collection.Value)
                Case TypeOf Me.UserInputComboBox.SelectedItem Is Computer
                    Dim computer As Computer = UserInputComboBox.SelectedItem
                    Me.ResourceExplorer.SelectedNode = LoadComputerNode(computer)
            End Select
        Else
            If Not String.IsNullOrWhiteSpace(Me.UserInputComboBox.Text) Then
                SearchUserInputText(Me.UserInputComboBox.Text)
            End If
        End If
    End Sub

    Private Sub NewButton_Click(sender As Object, e As EventArgs) Handles NewButton.Click
        Try
            Dim newInstanceArg = $"-{NameOf(Switch.NewInstance)}"
            Process.Start([Global].AppPath, newInstanceArg)
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub SearchUserInputText(userInputText As String)
        Dim searcher As New DataSourceSearcher(userInputText, Me.BindingSource)

        Dim searchResults As List(Of Computer) = searcher.GetComputers()
        If searchResults IsNot Nothing AndAlso searchResults.Count > 0 Then
            If searchResults.Count = 1 Then
                Me.ResourceExplorer.SelectedNode = LoadComputerNode(searchResults.Single())
            Else
                Me.ResourceExplorer.SelectedNode = LoadCollectionNode(userInputText, searcher)
            End If
        End If
    End Sub

    Private Sub GetInfo(sender As Object, e As EventArgs)
        Dim listView As ListView = Nothing
        Select Case True
            Case TypeOf sender Is ListView
                listView = sender
            Case TypeOf sender Is ToolStripMenuItem
                listView = sender.Tag
        End Select

        If listView.SelectedItems.Count > 0 Then
            Me.ResourceExplorer.SelectedNode = Me.ResourceExplorer.SelectedNode.Nodes(listView.SelectedItems(0).Name)
        End If
    End Sub

#End Region

#Region " TreeView "

    Private Sub ResourceExplorer_AfterSelect(sender As Object, e As TreeViewEventArgs) Handles ResourceExplorer.AfterSelect
        ' Clear work area
        Me.AcceptButton = SubmitButton
        ResetSplitContainer()

        ' If a root node is chosen discontinue processing
        If Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).IsSelected Then Exit Sub
        If Me.ResourceExplorer.Nodes(NameOf(Nodes.Settings)).IsSelected Then Exit Sub

        ' If a direct (non-recursive) root node child is selected, display a ListView, listing all the items underneath that node
        If Me.ResourceExplorer.SelectedNode.Parent.Name = NameOf(Nodes.RootNode) OrElse ResourceExplorer.SelectedNode.Parent.Name = NameOf(Nodes.Settings) Then DisplayChildItemControls()

        ' If a direct (non-recursive) computer node child is selected, load computer information utilizing the appropriate controller class
        If Me.ResourceExplorer.SelectedNode.Parent.Name = NameOf(Nodes.Computers) Then DisplayComputerControl()
        If Me.ResourceExplorer.SelectedNode.Parent.Name = NameOf(Nodes.Collections) Then DisplayCollectionControl(Nothing, Nothing)
        If Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Task)).IsSelected Then DisplayTaskControl()
        If Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Query)).IsSelected Then DisplayQueryControl()
        If Me.ResourceExplorer.SelectedNode.Parent.Name = NameOf(Nodes.Results) Then DisplayResultControl()
        If Me.ResourceExplorer.SelectedNode.Parent.Name = NameOf(Nodes.CustomActions) Then DisplayCustomActionControl()
    End Sub

    Private Sub ResourceExplorer_DragDrop(sender As Object, e As DragEventArgs) Handles ResourceExplorer.DragDrop
        MassImportFromDragAndDrop(sender, e)
    End Sub

    Private Sub ResourceExplorer_DragEnter(sender As Object, e As DragEventArgs) Handles ResourceExplorer.DragEnter
        If e.Data.GetDataPresent(DataFormats.Text) OrElse e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Function LoadComputerNode(computer As Computer, Optional isMassImport As Boolean = False) As TreeNode
        If computer Is Nothing Then Return Nothing

        ' Add the computer name to user settings for it to be restored on launch
        UserSettingsAddComputer(computer)

        Me.UserInputComboBox.SelectedItem = computer

        Dim computerNodes = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes

        Dim nodeMatch = computerNodes _
                            .OfType(Of TreeNode) _
                            .SingleOrDefault(Function(t) t.Name.ToUpper().Trim() = computer.Value.ToUpper().Trim())

        If nodeMatch IsNot Nothing Then
            ' This computer is already loaded
            If Not isMassImport Then
                ' Not currently importing so remove the node and re-load it below
                Me.ResourceExplorer.SelectedNode = nodeMatch
                Me.UIThread(Sub() nodeMatch.Remove())
            Else
                ' This computer is part of a mass import so do not re-load it
                Return nodeMatch
            End If
        End If

        ' Create a new node for this computer
        Dim newComputerNode As New TreeNode() With
        {
            .Name = computer.Value,
            .Text = computer.Display
        }

        Dim computerPanel As New ComputerPanel(computer, Me, newComputerNode)
        newComputerNode.Tag = computerPanel
        newComputerNode.ContextMenuStrip = computerPanel.NewComputerMenuStrip(ComputerPanel.ConnectionStatuses.Offline)

        Me.UIThread(Sub() computerNodes.Add(newComputerNode))

        Return newComputerNode
    End Function

    Private Function LoadCollectionNode(caption As String, searcher As DataSourceSearcher) As TreeNode
        caption = $"Search: {caption}"

        Dim collectionNodes = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Collections)).Nodes

        Dim nodeMatch = collectionNodes _
            .OfType(Of TreeNode) _
            .SingleOrDefault(Function(t) t.Text.ToUpper().Trim() = caption.ToUpper().Trim())

        If nodeMatch IsNot Nothing Then
            ' A collection with this search query is already loaded
            Return nodeMatch
        End If

        ' This is a new search query so create a new node for it
        Dim newCollectionNode As New TreeNode() With
        {
            .Text = caption,
            .Tag = searcher
        }

        collectionNodes.Add(newCollectionNode)

        Return newCollectionNode
    End Function

    Private Sub RemoveSelectedCollectionNode()
        Dim selectedNode = Me.ResourceExplorer.SelectedNode
        If selectedNode.Parent.Name = NameOf(Nodes.Collections) Then
            ' The currently selected node is indeed a collection
            Me.ResourceExplorer.Nodes.Remove(selectedNode)
        End If
    End Sub

    Private Sub RemoveAllComputerNodes()
        If Me.ImportWorker.IsBusy Then
            Dim message = "Computers are still being imported. Would you like to cancel the import and clear all results?"
            If MsgBox(message, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                Me.ImportWorker.CancelAsync()
            Else
                Exit Sub
            End If
        End If

        Dim computersNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers))
        RemoveMenuItem(computersNode, "Load All", 0)

        Dim garbageCollector As New GarbageCollector()
        For Each node As TreeNode In computersNode.Nodes
            Dim computerPanel As ComputerPanel = node.Tag
            garbageCollector.AddToGarbage = computerPanel
        Next
        garbageCollector.DisposeAsync()

        computersNode.Nodes.Clear()
        Me.ClearButton.PerformClick()

        UserSettingsClearComputers()
    End Sub

    Private Sub LoadAllComputerNodes()
        Dim computersNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers))

        RemoveMenuItem(computersNode, "Load All", 0)

        For Each node As TreeNode In computersNode.Nodes
            Dim computerPanel As ComputerPanel = node.Tag
            If Not computerPanel.IsLoaded Then
                If Not computerPanel.InitWorker.IsBusy Then
                    computerPanel.InitWorker.RunWorkerAsync()
                End If
            End If
        Next
    End Sub

    Private Sub DisplayChildItemControls()
        If Me.ResourceExplorer.SelectedNode Is Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)) Then
            RefreshComputerListView()
        Else
            Dim listView As ListView = NewComputerListView(Me.ResourceExplorer.SelectedNode.Text)
            If Me.ResourceExplorer.SelectedNode.Nodes.Count > 0 Then
                listView.BeginUpdate()
                For Index As Integer = 0 To Me.ResourceExplorer.SelectedNode.Nodes.Count - 1
                    listView.Items.Add(New ListViewItem(Me.ResourceExplorer.SelectedNode.Nodes(Index).Name) With
                    {
                        .Name = Me.ResourceExplorer.SelectedNode.Nodes(Index).Name,
                        .ForeColor = Me.ResourceExplorer.SelectedNode.Nodes(Index).ForeColor
                    })
                Next

                listView.EndUpdate()
            End If

            Me.MainSplitContainer.Panel2.InvokeAddControl(listView)
        End If
    End Sub

    Private Sub DisplayComputerControl()
        Dim panel As ComputerPanel = Me.ResourceExplorer.SelectedNode.Tag
        Me.UIThread(Sub() Me.UserInputComboBox.SelectedItem = panel.Computer)

        If Not panel.InitWorker.IsBusy Then
            SetDefaultProperties(panel)
            ResetSplitContainer()

            Me.MainSplitContainer.Panel2.InvokeAddControl(panel)

            If Not panel.Initialized Then
                panel.Timeout = 20
                panel.IsMassImport = False
                panel.InitWorker.RunWorkerAsync()
            End If
        End If
    End Sub

    Private Sub DisplayCollectionControl(sender As Object, e As EventArgs)
        Dim selectedNode = Me.ResourceExplorer.SelectedNode

        ' Display collection information when a child collection node is selected
        If selectedNode.Tag IsNot Nothing Then
            Dim collectionContext As CollectionControl = Nothing

            Select Case True
                Case TypeOf selectedNode.Tag Is String
                    collectionContext = New CollectionControl(Me, selectedNode.Tag.ToString(), Me.BindingSource.DataSource)
                Case TypeOf selectedNode.Tag Is DataSourceSearcher
                    Dim searcher As DataSourceSearcher = selectedNode.Tag
                    collectionContext = New CollectionControl(Me, selectedNode, searcher)
                    selectedNode.ContextMenuStrip = collectionContext.NewCollectionMenuStrip()
            End Select

            SetDefaultProperties(collectionContext)
            ResetSplitContainer()

            Me.MainSplitContainer.Panel2.InvokeAddControl(collectionContext)
            collectionContext.BeginEnumeration()
        Else
            Dim collectionContext As New CollectionControl(Me)
            SetDefaultProperties(collectionContext)
            ResetSplitContainer()

            Me.MainSplitContainer.Panel2.InvokeAddControl(collectionContext)
            collectionContext.BeginEnumeration()
        End If
    End Sub

    Private Sub DisplayTaskControl()
        Dim taskNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Task))

        If taskNode.Tag Is Nothing Then
            ResetSplitContainer()
            Dim task As New TaskControl(Me)
            SetDefaultProperties(task)

            Me.MainSplitContainer.Panel2.InvokeAddControl(task)
            taskNode.Tag = task
        Else
            ResetSplitContainer()
            Dim styling As TaskControl = CType(taskNode.Tag, TaskControl)
            styling.LoadCollectionInfo()
            MainSplitContainer.Panel2.InvokeAddControl(styling)
        End If
    End Sub

    Private Sub DisplayQueryControl()
        If Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Query)).Tag Is Nothing Then
            ResetSplitContainer()
            Dim query As New QueryControl(Me)
            SetDefaultProperties(query)

            Me.MainSplitContainer.Panel2.InvokeAddControl(query)
            Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Query)).Tag = query
        Else
            ResetSplitContainer()
            Dim query As QueryControl = CType(ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Query)).Tag, QueryControl)

            query.LoadCollectionInfo()
            Me.MainSplitContainer.Panel2.InvokeAddControl(query)
        End If
    End Sub

    Private Sub DisplayResultControl()
        ResetSplitContainer()

        Dim resultControl As Control = Me.ResourceExplorer.SelectedNode.Tag
        Me.MainSplitContainer.Panel2.InvokeAddControl(resultControl)
    End Sub

    Private Sub DisplayCustomActionControl()
        ResetSplitContainer()

        Dim actionControl As Control = Me.ResourceExplorer.SelectedNode.Tag
        Me.MainSplitContainer.Panel2.InvokeAddControl(actionControl)
    End Sub

    Private Sub SortComputerNodes()
        Me.ResourceExplorer.TreeViewNodeSorter = New Comparer(ResourceExplorer, True)
        Me.ResourceExplorer.Sort()
    End Sub

    Public Function IsNodeSelected(computerPanel As ComputerPanel) As Boolean
        Dim selected As Boolean = False

        Me.UIThread(Sub()
                        If Me.ResourceExplorer.SelectedNode.Tag Is computerPanel Then
                            selected = True
                        End If
                    End Sub)

        Return selected
    End Function

    Private Sub ExportSelectedComputerNodes()
        Try
            Dim saveDialog As New SaveFileDialog() With {.Filter = "Text File|*.txt", .OverwritePrompt = True}
            If saveDialog.ShowDialog() = Windows.Forms.DialogResult.OK Then
                Dim computerExportText As String = Nothing
                For Each treeNode As TreeNode In ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes
                    If computerExportText Is Nothing Then
                        computerExportText = treeNode.Name
                    Else
                        computerExportText += Environment.NewLine & treeNode.Name
                    End If
                Next

                If File.Exists(saveDialog.FileName) Then
                    File.Delete(saveDialog.FileName)
                End If

                File.WriteAllText(saveDialog.FileName, computerExportText)
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub RemoveComputerNodesByStatus(connectionStatuses As ComputerPanel.ConnectionStatuses())
        If Me.ImportWorker.IsBusy Then
            Dim message = "Computers are still being imported. Would you like to cancel the import and clear all results?"
            If MsgBox(message, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                ImportWorker.CancelAsync()
            Else
                Exit Sub
            End If
        End If

        Dim garbageCollector As New GarbageCollector()
        Dim index As Integer = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes.Count - 1
        For Each connectionStatus As ComputerPanel.ConnectionStatuses In connectionStatuses
            Do Until index = -1
                Dim node = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(index)
                Dim computerPanel As ComputerPanel = node.Tag

                If computerPanel.ConnectionStatus = connectionStatus Then
                    garbageCollector.AddToGarbage = computerPanel
                    SyncLock settingsLock
                        UserSettingsRemoveComputer(computerPanel.Computer)
                        node.Remove()
                    End SyncLock
                End If

                index -= 1
            Loop
        Next

        garbageCollector.DisposeAsync()
    End Sub

    Private Sub SetNodeText(node As TreeNode, text As String)
        Me.UIThread(Sub() node.Text = text)
    End Sub

#End Region

#Region " Context Menu Strips "

    Private Sub PopulateDynamicMenuItems(sender As Object, e As EventArgs)
        Dim menuStrip As ContextMenuStrip = sender

        Dim computersNodes = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers))
        If computersNodes.Nodes.Count > 0 Then
            Select Case menuStrip.SourceControl.Name

                Case Me.ResourceExplorer.Name
                    ' Remove "load all computers" menu item if all computers are already loaded
                    If computersNodes.Nodes.OfType(Of TreeNode)().All(Function(t) CType(t.Tag, ComputerPanel).IsLoaded) Then
                        RemoveMenuItem(computersNodes, "Load All", 0)
                    End If
                    With menuStrip.Items
                        .Add("Export Computers", Nothing, AddressOf ExportSelectedComputerNodes)
                        .Add(New ToolStripSeparator())
                        .Add(New ToolStripMenuItem("Remove Offline Computers", Nothing, Sub() RemoveComputerNodesByStatus({ComputerPanel.ConnectionStatuses.Offline})))
                        .Add(New ToolStripMenuItem("Remove Online Computers", Nothing, Sub() RemoveComputerNodesByStatus({ComputerPanel.ConnectionStatuses.Online, ComputerPanel.ConnectionStatuses.OnlineDegraded, ComputerPanel.ConnectionStatuses.OnlineSlow})))
                        .Add("Sort Computers", Nothing, AddressOf SortComputerNodes)
                    End With

                Case NameOf(Nodes.Computers)
                    Dim listViewComputer As ListView = menuStrip.SourceControl
                    If listViewComputer.SelectedItems.Count >= 1 Then
                        menuStrip.Items.Add("Export Selected Computers", Nothing, AddressOf ExportSelectedComputerNodes)
                    Else
                        With menuStrip.Items
                            .Add("Export Computers", Nothing, AddressOf ExportSelectedComputerNodes)
                            .Add(New ToolStripSeparator())
                            .Add(New ToolStripMenuItem("Remove Offline Computers", Nothing, Sub() RemoveComputerNodesByStatus({ComputerPanel.ConnectionStatuses.Offline})))
                            .Add(New ToolStripMenuItem("Remove Online Computers", Nothing, Sub() RemoveComputerNodesByStatus({ComputerPanel.ConnectionStatuses.Online, ComputerPanel.ConnectionStatuses.OnlineDegraded, ComputerPanel.ConnectionStatuses.OnlineSlow})))
                        End With
                    End If
            End Select
        End If
    End Sub

    Private Sub RemoveDynamicMenuItems(sender As Object, e As EventArgs)
        Dim menuStrip As ContextMenuStrip = sender
        If menuStrip.Items.Count > 3 Then
            For i = 3 To menuStrip.Items.Count - 1
                menuStrip.Items.RemoveAt(3)
            Next
        End If
    End Sub

    Private Sub NewComputerMenuStrip(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then

            Dim listView As ListView = sender
            Dim computersNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers))
            Dim menuStrip As ContextMenuStrip = computersNode.ContextMenuStrip

            Select Case listView.SelectedItems.Count

                Case Is > 1
                    With menuStrip.Items
                        .Add(New ToolStripSeparator())
                        .Add(New ToolStripMenuItem("Remove Selected Computers", Nothing, Sub() RemoveSelectedComputerNodes(listView)))
                        .Add(New ToolStripSeparator())
                        .Add(New ToolStripMenuItem("Logoff Selected Computers", Nothing, Sub() LogOffSelectedComputerNodes(listView)))
                        .Add(New ToolStripMenuItem("Reboot Selected Computers", Nothing, Sub() RebootSelectedComputerNodes(listView)))
                        .Add(New ToolStripSeparator())
                        .Add(New ToolStripMenuItem("Enable Computers", Nothing, Sub() EnableSelectedComputerNodes(listView)))
                        .Add(New ToolStripMenuItem("Disable Computers", Nothing, Sub() DisableSelectedComputerNodes(listView)))
                        .Add(New ToolStripSeparator())
                    End With

                    For Each action In New RegistryController().GetKeyValues(My.Settings.RegistryPathCustomActions, RegistryHive.CurrentUser)
                        Dim menuItem As New ToolStripMenuItem(action, Nothing, Sub() PerformCustomActionOnSelectedComputerNodes(listView, menuItem, e)) With {.Tag = New String() {action, "CUSTOM"}}
                        menuStrip.Items.Add(menuItem)
                    Next

                Case 1
                    Dim computerPanel As ComputerPanel = computersNode.Nodes(listView.SelectedItems(0).Name).Tag
                    menuStrip = computerPanel.NewComputerMenuStrip(computerPanel.ConnectionStatus)
                    Dim getInfoMenuItem As New ToolStripMenuItem("Get Info", Nothing, AddressOf GetInfo)
                    getInfoMenuItem.Tag = listView
                    menuStrip.Items.Insert(0, getInfoMenuItem)

            End Select

            listView.ContextMenuStrip = menuStrip
        End If
    End Sub

    Private Sub RemoveMenuItem(node As TreeNode, itemText As String, index As Integer)
        Dim menuStrip = node.ContextMenuStrip

        If menuStrip.Items(index).Text = itemText Then
            menuStrip.Items.RemoveAt(0)
        End If
    End Sub

#End Region

#Region " Collection Management "

    Private Sub NewCollection()
        Dim registry As New RegistryController()
        Dim collectionNameForm As New CollectionNameForm()

        Dim columnOptionsMenuStrip As New ContextMenuStrip()
        columnOptionsMenuStrip.Items.Add("Delete", Nothing, AddressOf DeleteCollection)

        If collectionNameForm.ShowDialog() = DialogResult.OK Then
            Dim treeNode As New TreeNode() With
                {
                    .Name = collectionNameForm.ColumnNameTextBox.Text,
                    .Text = collectionNameForm.ColumnNameTextBox.Text,
                    .ContextMenuStrip = columnOptionsMenuStrip
                }

            Dim collectionKey As String = Path.Combine(My.Settings.RegistryPathCollections, collectionNameForm.ColumnNameTextBox.Text)
            registry.NewKey(collectionKey, RegistryHive.CurrentUser)

            Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Collections)).Nodes.Add(treeNode)
            Me.ResourceExplorer.SelectedNode = treeNode
        End If
    End Sub

    Private Sub DeleteCollection()
        Dim collectionKey As String = Path.Combine(My.Settings.RegistryPathCollections, ResourceExplorer.SelectedNode.Text)
        Dim registry As New RegistryController()
        registry.DeleteKey(collectionKey, RegistryHive.CurrentUser)

        Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Collections)).Nodes.Remove(ResourceExplorer.SelectedNode)
    End Sub

#End Region

#Region " Computer ListView "

    Public Sub RefreshComputerListView()
        Dim rootComputerTreeNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers))

        If Me.ResourceExplorer.SelectedNode Is rootComputerTreeNode Then
            If Me.MainSplitContainer.Panel2.Controls.Count > 0 Then
                Dim computerListView As ListView = Me.MainSplitContainer.Panel2.Controls(0)
                If computerListView.Name = NameOf(Nodes.Computers) Then

                    Dim treeNodes As TreeNodeCollection = rootComputerTreeNode.Nodes
                    If Not computerListView.Items.Cast(Of ListViewItem)().Any(Function(listViewItem) treeNodes.ContainsKey(listViewItem.Name)) Then

                        If Not treeNodes.Cast(Of TreeNode)().Any(Function(treeNode) Not computerListView.Items.ContainsKey(treeNode.Name)) Then
                            RefreshSingleComputerNode(computerListView)
                        End If

                    End If
                End If
            End If

            RefreshAllComputerNodes()
        End If
    End Sub

    Private Sub RefreshSingleComputerNode(listViewComputer As ListView)
        listViewComputer.BeginUpdate()

        For Each computerListViewItem In listViewComputer.Items
            Dim computerNode As TreeNode = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name)

            computerListViewItem.ForeColor = computerNode.ForeColor
            computerListViewItem.Font = computerNode.NodeFont
        Next

        listViewComputer.EndUpdate()
    End Sub

    Private Sub RefreshAllComputerNodes()
        Dim selectedNode = Me.ResourceExplorer.SelectedNode

        Dim computerListView As ListView = NewComputerListView(selectedNode.Text)
        With computerListView
            .Name = NameOf(Nodes.Computers)
            .MultiSelect = True
            .AllowDrop = True
            .MultiSelect = True
        End With

        Me.MainSplitContainer.Panel2.InvokeClearControls()

        AddHandler computerListView.MouseUp, AddressOf NewComputerMenuStrip
        For index As Integer = 0 To (selectedNode.Nodes.Count - 1)
            Dim computerListViewItem As New ListViewItem(selectedNode.Nodes(index).Name) With
                {
                    .Name = selectedNode.Nodes(index).Name,
                    .ForeColor = selectedNode.Nodes(index).ForeColor,
                    .Font = selectedNode.Nodes(index).NodeFont
                }

            ' Computer node alterations
            If selectedNode.Name = NameOf(Nodes.Computers) Then
                computerListViewItem.Text = selectedNode.Nodes(index).Text
            End If

            computerListView.Items.Add(computerListViewItem)
        Next

        AddHandler computerListView.MouseDown, Sub() Me.MouseDownOnListView = True
        AddHandler computerListView.MouseUp, Sub() Me.MouseDownOnListView = False
        AddHandler computerListView.MouseMove, Sub() ComputerListView_MouseMove(computerListView)

        Me.MainSplitContainer.Panel2.InvokeAddControl(computerListView)
    End Sub

    Private Sub ComputerListView_MouseMove(computerListView As ListView)
        If Me.MouseDownOnListView Then
            If computerListView.SelectedItems.Count > 0 Then

                Dim copyString As String = Nothing
                For Each listViewItem As ListViewItem In computerListView.SelectedItems
                    If copyString Is Nothing Then
                        copyString = listViewItem.Name
                    Else
                        copyString += Environment.NewLine & listViewItem.Name
                    End If
                Next

                computerListView.DoDragDrop(copyString, DragDropEffects.Copy)
            End If
        End If
    End Sub

    Private Sub RemoveSelectedComputerNodes(computerListView As ListView)
        Dim garbageCollector As New GarbageCollector()

        For Each computerListViewItem As ListViewItem In computerListView.SelectedItems
            Dim computerPanel As ComputerPanel = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name).Tag
            garbageCollector.AddToGarbage = computerPanel

            SyncLock nodeLock
                ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name).Remove()
            End SyncLock
        Next

        garbageCollector.DisposeAsync()
        RefreshComputerListView()
    End Sub

    Private Sub LogOffSelectedComputerNodes(computerListView As ListView)
        For Each computerListViewItem As ListViewItem In computerListView.SelectedItems
            Dim computerPanel As ComputerPanel = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name).Tag

            If computerPanel.ConnectionStatus = ComputerPanel.ConnectionStatuses.Online Then
                computerPanel.InitiateLogoff()
            End If
        Next
    End Sub

    Private Sub RebootSelectedComputerNodes(computerListView As ListView)
        For Each computerListViewItem As ListViewItem In computerListView.SelectedItems
            Dim computerPanel As ComputerPanel = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name).Tag

            If computerPanel.ConnectionStatus = ComputerPanel.ConnectionStatuses.Online Then
                computerPanel.InitiateNoPromptRestart()
            End If
        Next
    End Sub

    Private Sub EnableSelectedComputerNodes(computerListView As ListView)
        For Each computerListViewItem As ListViewItem In computerListView.SelectedItems
            Dim computerPanel As ComputerPanel = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name).Tag
            computerPanel.InitiateEnableComputers()
        Next
    End Sub

    Private Sub DisableSelectedComputerNodes(computerListView As ListView)
        For Each computerListViewItem As ListViewItem In computerListView.SelectedItems
            Dim computerPanel As ComputerPanel = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(computerListViewItem.Name).Tag
            computerPanel.InitiateDisableComputers()
        Next
    End Sub

    Private Sub PerformCustomActionOnSelectedComputerNodes(listView As ListView, customActionMenuItem As ToolStripMenuItem, e As MouseEventArgs)
        [Global].SetPsExecBinaryPath()

        For Each listViewItem As ListViewItem In listView.SelectedItems
            Dim computerPanel As ComputerPanel = ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes(listViewItem.Name).Tag

            If computerPanel.ConnectionStatus <> ComputerPanel.ConnectionStatuses.Offline Then
                computerPanel.InitiateCustomAction(customActionMenuItem, e)
            End If
        Next
    End Sub

#End Region

#Region " Control Manipulation "

    Private Sub SetDefaultProperties(ByRef control As Control)
        With control
            .Height = Me.MainSplitContainer.Panel2.Height
            .Width = Me.MainSplitContainer.Panel2.Width
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Right
        End With
    End Sub

    Private Sub SetMainWindowTitle()
        Dim newTitle As String = InputBox("Window Title Change Prompt", "New Title", Me.DefaultTitle)

        If Not String.IsNullOrWhiteSpace(newTitle) Then
            If newTitle.Length > 50 Then
                MsgBox("The window title is too long. Max: 50 Characters")
                SetMainWindowTitle()
                Exit Sub
            End If

            Me.Text = $"{newTitle} - {Me.DefaultTitle}"
        End If
    End Sub

    Private Sub SetAcceptButton(button As Button)
        Me.AcceptButton = button
    End Sub

    Private Function ComboBoxFocused() As Boolean
        Return Me.UserInputComboBox.Focused
    End Function

    Private Function NewComputerListView(name As String) As ListView
        ' If a child node of the RootNode is selected, this creates a ListView on the fly to display subitems contained within the node selected.
        Dim listView As New ListView()
        With listView
            .Anchor = AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Top + AnchorStyles.Bottom
            .View = View.Details
            .HideSelection = False
            .Width = Me.MainSplitContainer.Panel2.Width
            .Height = Me.MainSplitContainer.Panel2.Height
            .MultiSelect = False
            .Columns.Add(name)
            .Columns(0).Width = listView.Width - 5
        End With

        AddHandler listView.DoubleClick, AddressOf GetInfo

        Return listView
    End Function

    Public Sub ResetSplitContainer()
        Me.UIThread(Sub()
                        With Me.MainSplitContainer.Panel2
                            .BackgroundImage = Nothing
                            .BackColor = Color.Transparent
                            .Controls.Clear()
                        End With
                    End Sub)
    End Sub

    Private Sub PopulateCollectionOptions(collectionOptions As CollectionOptionsControl)
        collectionOptions.ComputerStatusListView.Items.Clear()

        Dim registry As New RegistryController
        Dim computerNames As New List(Of String)

        Dim collectionsRegistryPath As String = Path.Combine(My.Settings.RegistryPathCollections, Me.ResourceExplorer.SelectedNode.Text)
        If registry.GetKeyValues(collectionsRegistryPath, RegistryHive.CurrentUser) IsNot Nothing Then
            For Each computerName As String In registry.GetKeyValues(collectionsRegistryPath, RegistryHive.CurrentUser)
                computerNames.Add(computerName)
            Next
        End If

        collectionOptions.AddToCollectionListView(computerNames)
    End Sub

    Private Sub UserInputComboBox_DragDrop(sender As Object, e As DragEventArgs) Handles UserInputComboBox.DragDrop
        MassImportFromDragAndDrop(sender, e)
    End Sub

    Private Sub UserInputComboBox_DragEnter(sender As Object, e As DragEventArgs) Handles UserInputComboBox.DragEnter
        If e.Data.GetDataPresent(DataFormats.Text) Then
            e.Effect = DragDropEffects.All
        End If

        If e.Data.GetDataPresent(DataFormats.FileDrop) Then
            e.Effect = DragDropEffects.All
        End If
    End Sub

    Private Sub UserInputComboBox_KeyDown(sender As Object, e As KeyEventArgs) Handles UserInputComboBox.KeyDown
        If e.Modifiers = Keys.Control AndAlso e.KeyCode = Keys.V Then
            If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
                Dim clipboardText As String = Clipboard.GetText().Trim()
                If clipboardText.Split(" ").Count() > 1 OrElse clipboardText.Split(ControlChars.Tab).Count() > 1 OrElse clipboardText.Split(Environment.NewLine).Count() > 1 Then
                    MassImportFromClipboard()
                End If
            End If

            e.Handled = True
        End If
    End Sub

#End Region

#Region " Custom Action Management "

    Private Sub NewCustomAction()
        ResetSplitContainer()
        Me.MainSplitContainer.Panel2.InvokeAddControl(New CustomActionsControl(Me))
    End Sub

    Private Sub DeleteCustomAction(ByRef treeNode As TreeNode, customAction As String)
        Dim customActionKey As String = Path.Combine(My.Settings.RegistryPathCustomActions, customAction)
        Dim registry As New RegistryController()
        registry.DeleteKey(customActionKey, RegistryHive.CurrentUser)

        treeNode.Remove()
    End Sub

#End Region

#Region " Mass Import "

    Private Sub MassImportFromClipboard()
        If Me.ImportWorker.IsBusy Then
            Dim message = "The system tool is currently processing an import job. Would you like to cancel the import job and begin a new one?"
            If MsgBox(message, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                Me.ImportWorker.CancelAsync()
            End If
        End If

        If Clipboard.GetDataObject.GetDataPresent(DataFormats.Text) Then
            If Not Me.ImportWorker.IsBusy Then
                Me.ImportWorker.RunWorkerAsync(Clipboard.GetText)
            End If
        End If
    End Sub

    Private Sub MassImportFromDragAndDrop(sender As Object, e As DragEventArgs)
        Try
            If e.Data.GetDataPresent(DataFormats.Text) Then
                If Me.ImportWorker.IsBusy Then
                    Dim message = "The system tool is currently processing an import job. Would you like to cancel the import job and begin a new one?"
                    If MsgBox(message, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                        Me.ImportWorker.CancelAsync()
                    End If
                End If

                If Not Me.ImportWorker.IsBusy Then
                    Me.ImportWorker.RunWorkerAsync(e.Data.GetData(DataFormats.Text))
                End If

            ElseIf e.Data.GetDataPresent(DataFormats.FileDrop) Then
                Dim stringBuilder As StringBuilder = Nothing

                Dim droppedFiles As String() = e.Data.GetData(DataFormats.FileDrop)
                For Each file As String In droppedFiles
                    If Path.GetExtension(file) = ".txt" Then
                        stringBuilder.Append(ControlChars.Tab & IO.File.ReadAllText(file))
                    End If
                Next

                If stringBuilder IsNot Nothing AndAlso stringBuilder.Length > 0 AndAlso Not Me.ImportWorker.IsBusy Then
                    Me.ImportWorker.RunWorkerAsync(stringBuilder.ToString())
                End If
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub MassImportFromTextFile()
        Try
            If Me.ImportWorker.IsBusy Then
                Dim message = "The system tool is currently processing an import job. Would you like to cancel the import job and begin a new one?"
                If MsgBox(message, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = MsgBoxResult.Yes Then
                    Me.ImportWorker.CancelAsync()
                Else
                    Exit Sub
                End If
            End If

            Dim openDialog As New OpenFileDialog() With
                {
                    .Filter = "Text Files|*.txt",
                    .Multiselect = True,
                    .Title = "Select file(s) to import"
                }

            If openDialog.ShowDialog() = DialogResult.OK Then
                Dim stringBuilder As StringBuilder = Nothing
                For Each file As String In openDialog.FileNames
                    stringBuilder.Append(ControlChars.Tab & IO.File.ReadAllText(file))
                Next

                If stringBuilder IsNot Nothing AndAlso stringBuilder.Length > 0 AndAlso Not Me.ImportWorker.IsBusy Then
                    Me.ImportWorker.RunWorkerAsync(stringBuilder)
                End If
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub MassImportProcessComputers(sender As BackgroundWorker, e As DoWorkEventArgs)
        Dim importText As String = CType(e.Argument, String)
        If Not String.IsNullOrWhiteSpace(importText) Then

            Dim currentComputerTreeNodes As New List(Of TreeNode)
            currentComputerTreeNodes.AddRange(importText.Split(Environment.NewLine).Select(Function(ComputerName) MassImportNewComputerTreeNode(ComputerName)))

            If currentComputerTreeNodes IsNot Nothing Then
                For Each treeNode As TreeNode In currentComputerTreeNodes
                    If sender.CancellationPending Then Exit Sub
                    If treeNode IsNot Nothing Then
                        Dim worker As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
                        AddHandler worker.DoWork, AddressOf MassImportTreeNodeLoader
                        worker.RunWorkerAsync(treeNode)
                    End If
                Next
            End If

        End If
    End Sub

    Private Sub MassImportTreeNodeLoader(sender As Object, e As DoWorkEventArgs)
        Dim panel As ComputerPanel = CType(e.Argument, TreeNode).Tag

        If Not panel.InitWorker.IsBusy AndAlso Not panel.Initialized Then
            With panel
                .Timeout = 5
                .IsMassImport = True
                .UseRandomTimeoutForMassImport = True
                .StartLoadingComputer()
            End With
        End If
    End Sub

    Private Function MassImportNewComputerTreeNode(searchText As String) As TreeNode
        If Not String.IsNullOrWhiteSpace(searchText) AndAlso Not searchText.Length < 5 AndAlso Not searchText.Length > 20 Then
            Dim searcher As New DataSourceSearcher(searchText, Me.BindingSource)

            Dim computer As Computer = searcher.GetComputer()
            If computer IsNot Nothing Then
                Dim node As TreeNode = LoadComputerNode(computer, True)
                Return node
            End If
        End If

        Return Nothing
    End Function

#End Region

#Region " Ping Monitor "

    Private Async Sub RunPingMonitor()
        ' Start the ping monitor thread
        Dim cts As New CancellationTokenSource()
        Dim token As CancellationToken = cts.Token

        AddHandler Me.FormClosing, AddressOf cts.Cancel

        While Not cts.IsCancellationRequested
            Await Task.Run(Sub() PingComputerNodes(token)).ConfigureAwait(False)
        End While
    End Sub

    Private Sub PingComputerNodes(token As CancellationToken)
        ' *** This method runs continuously on the thread pool until cancelled on the Me.FormClosing event ***

        ' Wait for 10 seconds before the next round of pinging
        Thread.Sleep(10000)

        Try
            ' Do not ping computers during import
            If Not Me.ImportWorker.IsBusy Then
                ' Get a collection of any loaded computer nodes
                Dim computerNodes = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes
                For Each node As TreeNode In computerNodes
                    token.ThrowIfCancellationRequested()
                    PingNode(node)
                Next
            End If

        Catch ex As OperationCanceledException
            ' Operation was cancelled
            Exit Sub
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub PingNode(node As TreeNode)
        Dim computerPanel As ComputerPanel = node.Tag
        If computerPanel.IsLoaded Then
            computerPanel.ReportConnectionStatus()
        End If
    End Sub

#End Region

#Region " LDAP Monitor "

    Private Sub UpdateLDAPDataBindings(sender As BackgroundWorker, e As DoWorkEventArgs)
        ' *** This method runs continuously on the thread pool until cancelled on the Me.FormClosing event ***
        SyncLock ldapLock

            ' Check for newly added computers in ldap every 20 seconds
            Thread.Sleep(20000)

            Try
                ' Query ldap for a list of all computers that are either new or updated since the last incremental update (or full load)
                Using searchResults As SearchResultCollection = LDAPSearcher.GetIncrementalUpdateResults(Me.LastChangedLDAPTime)
                    If searchResults?.Count > 0 Then
                        ' At least one computer has been either updated or added to ldap so first set the new last changed time for the next incremental update
                        Me.LastChangedLDAPTime = LDAPSearcher.GetLastChangedTime(searchResults)

                        ' Begin looping through all of the results returned by the incremental update ldap query
                        For Each result As SearchResult In searchResults
                            If sender.CancellationPending Then
                                e.Cancel = True
                                Return
                            End If

                            Dim ldapComputer As New Computer(result)
                            Dim dataBoundComputer = Me.BindingSource.OfType(Of Computer).SingleOrDefault(Function(c) c.Value = ldapComputer.Value)
                            Dim index As Integer = Me.BindingSource.IndexOf(dataBoundComputer)
                            If index <> -1 Then
                                ' The computer name was found in both ldap and the bound data source
                                ' Get the collection of loaded computer nodes
                                Dim computerNodes = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes.Cast(Of TreeNode)()

                                ' See if this computer has been loaded to a node
                                Dim dataBoundNode = computerNodes.SingleOrDefault(Function(t) t.Tag.Computer.Value = dataBoundComputer.Value)
                                If dataBoundNode IsNot Nothing Then
                                    ' Computer is loaded to a node, so update it
                                    dataBoundNode.Tag.Computer = ldapComputer
                                    SetNodeText(dataBoundNode, ldapComputer.Display)
                                End If

                                ' Also update the computer object in the binding source
                                Me.BindingSource(index) = ldapComputer
                            Else
                                ' The computer is in ldap but not in the binding source, so add it
                                Me.BindingSource.Add(ldapComputer)
                                index = Me.BindingSource.IndexOf(ldapComputer)
                            End If

                            ResetLDAPDataBindings(index)
                        Next
                    End If
                End Using
            Catch ex As Exception
                LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
            End Try
        End SyncLock
    End Sub

    Private Sub ResetLDAPDataBindings(index As Integer)
        If index = -1 Then Return
        Me.UIThread(Sub() Me.BindingSource.ResetItem(index))
    End Sub

#End Region

#Region " Status Strip Monitor "

    Private Async Sub RunStatusStripMonitor()
        ' Prepare the status strip
        SetStatusStripAppearance()
        AddVersionStatusLabel()

        ' Create performance monitors
        Dim proc As Process = Process.GetCurrentProcess()
        Dim cpuCounter As PerformanceCounter = Await Task.Run(Function() NewPerformanceCounter(proc, "% Processor Time"))
        Dim memCounter As PerformanceCounter = Await Task.Run(Function() NewPerformanceCounter(proc, "Working Set - Private"))

        ' Start the status strip monitor thread
        Dim cts As New CancellationTokenSource()
        Dim token As CancellationToken = cts.Token
        AddHandler Me.FormClosing, AddressOf cts.Cancel
        While Not cts.IsCancellationRequested
            Await Task.Run(Sub() UpdateStatusStrip(token, proc, cpuCounter, memCounter))
        End While
    End Sub

    Private Sub SetStatusStripAppearance()
        ' Configure appearance of the status strip labels
        For Each statusLabel In StatusStrip1.Items.OfType(Of ToolStripStatusLabel)
            statusLabel.Visible = False
            statusLabel.Text = String.Empty
            statusLabel.Alignment = ToolStripItemAlignment.Right

            ' Set the left hand border for all labels except the last one
            If StatusStrip1.Items.IndexOf(statusLabel) < StatusStrip1.Items.Count - 1 Then
                statusLabel.BorderSides = ToolStripStatusLabelBorderSides.Left
            End If
        Next
    End Sub

    Private Sub AddVersionStatusLabel()
        Dim versionLabel As New ToolStripStatusLabel()
        versionLabel.Text = $"Version: {[Global].AppVersion}"
        versionLabel.Alignment = ToolStripItemAlignment.Left

        StatusStrip1.Items.Add(versionLabel)
    End Sub

    Private Function NewPerformanceCounter(proc As Process, counterName As String) As PerformanceCounter
        Dim category As New PerformanceCounterCategory("Process")
        Dim instanceNames As String() = category.GetInstanceNames() _
            .Where(Function(i) i.StartsWith(proc.ProcessName)) _
            .ToArray()

        Dim pid As Integer = proc.Id

        For Each instanceName In instanceNames
            Using counter As New PerformanceCounter("Process", "ID Process", instanceName, True)
                If counter.RawValue <> Nothing AndAlso CInt(counter.RawValue) = proc.Id Then
                    counter.CounterName = counterName

                    Return counter
                End If
            End Using
        Next

        Return Nothing
    End Function

    Private Sub UpdateStatusStrip(token As CancellationToken, proc As Process, cpuCounter As PerformanceCounter, memCounter As PerformanceCounter)
        ' *** This method runs continuously on the thread pool until cancelled on the Me.FormClosing event ***
        SyncLock statusStripLock

            ' Update the status strip every 1 second
            Thread.Sleep(1000)
            Try
                token.ThrowIfCancellationRequested()

                ' Date time
                Dim dateTime As Date = Date.Now
                Dim dateTimeText As String = dateTime.ToLocalTime()

                ' Up time
                Dim upTime As TimeSpan = Date.Now - proc.StartTime
                Dim upTimeText = $"Up Time: {upTime.Days}:{upTime.Hours:D2}:{upTime.Minutes:D2}:{upTime.Seconds:D2}"

                ' CPU usage
                Dim cpuUsage As Integer = cpuCounter.NextValue() / Environment.ProcessorCount
                Dim cpuUsageText = $"CPU Usage: {cpuUsage}"

                ' Memory usage
                Dim memUsage As Single = memCounter.NextValue() / 1024 / 1024
                Dim memUsageText = $"Memory Usage: {memUsage:F1} MB"

                ' Connections count
                Dim connCount As Integer = GetConnectionsCount()
                Dim connCountText = $"Connections: {connCount}"

                ' Update the status labels display text
                ShowStatusStripItems()
                Me.UIThread(Sub()
                                Me.DateTimeStatusLabel.Text = dateTimeText
                                Me.UpTimeStatusLabel.Text = upTimeText
                                Me.CpuUsageStatusLabel.Text = cpuUsageText
                                Me.MemoryUsageStatusLabel.Text = memUsageText
                                Me.ConnectionsStatusLabel.Text = connCountText
                            End Sub)
            Catch ex As OperationCanceledException
                ' Operation was cancelled
                Exit Sub
            End Try
        End SyncLock
    End Sub

    Private Function GetConnectionsCount() As Integer
        Dim connnectionsCount As Integer = 0
        Me.ResourceExplorer.UIThread(Sub()
                                         Dim computerNodes = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Nodes
                                         For Each node As TreeNode In computerNodes
                                             Dim computerPanel As ComputerPanel = node.Tag
                                             If computerPanel.IsLoaded AndAlso computerPanel.ConnectionStatus <> ComputerPanel.ConnectionStatuses.Offline Then
                                                 connnectionsCount += 1
                                             End If
                                         Next
                                     End Sub)

        Return connnectionsCount
    End Function

    Private Sub ShowStatusStripItems()
        Me.UIThread(Sub()
                        For Each item As ToolStripItem In Me.StatusStrip1.Items
                            item.Visible = True
                        Next
                    End Sub)
    End Sub

#End Region

#Region " Update Monitor "

    Private Async Sub RunUpdateMonitor()
        ' Don't run update monitor when current version is unavailable
        If Not [Global].AppVersion.Equals("not found") Then
            ' The NuGet URI for the package feed
            Dim apiUri As New Uri(My.Settings.NuGetURI)

            ' Start the application update monitor thread
            Dim cts As New CancellationTokenSource()
            Dim token As CancellationToken = cts.Token
            AddHandler Me.FormClosing, AddressOf cts.Cancel
            While Not cts.IsCancellationRequested
                Await Task.Run(Sub() CheckForAppUpdate(token, apiUri, pkgName:=My.Settings.NuGetPkgName))
            End While
        End If
    End Sub

    Private Sub CheckForAppUpdate(token As CancellationToken, apiUri As Uri, pkgName As String)
        ' *** This method runs continuously on the thread pool until cancelled on the Me.FormClosing event ***
        SyncLock updateLock
            Try
                token.ThrowIfCancellationRequested()

                Me.AppUpdates = GetNuGetQueryResult(apiUri, pkgName, prerelease:=True)
                Dim latestVersion As String = Me.AppUpdates.Data.First().Version.Trim()

                If Not latestVersion.Equals([Global].AppVersion) Then
                    SetNewVersionAvailable(latestVersion)
                End If
            Catch ex As OperationCanceledException
                ' Operation was cancelled
                Exit Sub
            Catch ex As Exception
                LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
            End Try

            ' Check for update every 5 minutes
            Thread.Sleep(300000)
        End SyncLock
    End Sub

    Private Sub SetNewVersionAvailable(version As String)
        ' Alert the user that there is a new version available
        Me.Text = Me.DefaultTitle & " - New Version Available: " & version
        If Not Me.AppUpdateNotificationSent Then
            ' Only send notification once
            Me.UIThread(Sub() Me.FlashNotification(count:=3))
            Me.AppUpdateNotificationSent = True
        End If

        ' Set the registry entry to trigger the upgrade on next launch
        Dim registry As New RegistryController()
        Dim regEntry As String = NameOf(RegistryController.RegistrySettings.UpdateAvailable)
        registry.SetKeyValue(My.Settings.RegistryPath, regEntry, "1", hive:=RegistryHive.LocalMachine)
    End Sub

    Private Function GetNuGetQueryResult(apiUri As Uri, pkgName As String, Optional prerelease As Boolean = False) As NuGetSearchResult
        Dim result As NuGetSearchResult = Nothing
        Dim uri As New Uri(apiUri, "query2?q=" & pkgName)

        If prerelease Then
            Dim uriBuilder As New UriBuilder(uri.AbsoluteUri)
            Dim query = HttpUtility.ParseQueryString(uriBuilder.Query)
            query("prerelease") = "true"
            uriBuilder.Query = query.ToString()
            uri = uriBuilder.Uri
        End If

        Dim request As HttpWebRequest = WebRequest.Create(uri)
        request.UseDefaultCredentials = True

        Using stream = request.GetResponse().GetResponseStream()
            Using reader = New StreamReader(stream)
                Dim jsonText = reader.ReadToEnd()
                result = JsonConvert.DeserializeObject(Of NuGetSearchResult)(jsonText)
            End Using
        End Using

        Return result
    End Function

#End Region

#Region " User Settings "

    Private Sub UserSettingsLoadFormSize()
        Me.Text = DefaultTitle
        Me.Height = My.Settings.MainFormHeight
        Me.Width = My.Settings.MainFormWidth
    End Sub

    Private Sub UserSettingsSaveFormSize()
        Try
            With My.Settings
                .MainFormHeight = Me.Size.Height
                .MainFormWidth = Me.Size.Width
                .Save()
            End With
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub UserSettingsAddComputer(computer As Computer)
        If computer Is Nothing Then Return
        Try
            If My.Settings.ActiveComputers Is Nothing Then
                My.Settings.ActiveComputers = New Specialized.StringCollection()
            End If

            My.Settings.ActiveComputers.Add(computer.Value)
            My.Settings.Save()
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Sub UserSettingsRemoveComputer(computer As Computer)
        If computer Is Nothing Then Return
        Try
            My.Settings.ActiveComputers.Remove(computer.Value)
            My.Settings.Save()
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub UserSettingsClearComputers()
        Try
            My.Settings.ActiveComputers.Clear()
            My.Settings.Save()
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub UserSettingsLoadComputers()
        Try
            ' Don't load stored computers if the new instance switch was passed at launch
            Dim args = Environment.GetCommandLineArgs().Select(Function(a) a.ToUpper().Trim())
            Dim newInstanceArg = $"-{NameOf(Switch.NewInstance).ToUpper()}"

            If Not args.Contains(newInstanceArg) Then
                If My.Settings.ActiveComputers IsNot Nothing Then
                    ' Make a copy of the computer names to load
                    Dim computerNamesToLoad = My.Settings.ActiveComputers _
                        .Cast(Of String)() _
                        .ToArray()

                    ' This is the bound data loaded from ldap during LoadForm
                    Dim boundComputers = Me.BindingSource.OfType(Of Computer)

                    For Each computerName As String In computerNamesToLoad
                        ' Lookup the computer from the bound data
                        Dim computer = boundComputers.SingleOrDefault(Function(c) c.Value = computerName)

                        ' Remove the computer name from user settings
                        UserSettingsRemoveComputer(computer)

                        ' Load the computer node in the resource explorer
                        LoadComputerNode(computer)
                    Next

                    ' Once done loading all saved computers, add a "load all computers" menu item
                    Dim loadMenuItem As New ToolStripMenuItem("Load All", Nothing, AddressOf LoadAllComputerNodes)
                    Dim computersNode = Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers))
                    computersNode.ContextMenuStrip.Items.Insert(0, loadMenuItem)

                    ' And expand the computers node
                    Me.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Computers)).Expand()
                End If
            End If

        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

    End Sub

#End Region

End Class