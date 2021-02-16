Imports Microsoft.Win32

Public Class TaskControl

    Public Property OwnerForm As MainForm

    Public Sub New(ownerForm As MainForm)
        InitializeComponent()
        Me.OwnerForm = ownerForm
    End Sub

    Private Sub TaskControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCollectionInfo()
        If My.Settings.RecentTasks IsNot Nothing Then
            If My.Settings.RecentTasks.Count > 0 Then
                For Each recentTask As String In My.Settings.RecentTasks
                    Me.FileComboBox.Items.Add(recentTask)
                Next
            End If
        End If
    End Sub

    Public Sub LoadCollectionInfo()
        Me.CollectionsTreeView.Nodes.Clear()

        Dim collectionNodeMenuStrip As New ContextMenuStrip()
        collectionNodeMenuStrip.Items.Add("Get Info", Nothing, AddressOf GetInfo)

        Me.CollectionsTreeView.Nodes.Add(New TreeNode With
        {
            .Name = NameOf(Nodes.RootNode),
            .Text = NameOf(Nodes.Collections)
        })

        Dim registry As New RegistryController()
        Dim keyValues = registry.GetKeyValues(My.Settings.RegistryPathCollections, RegistryHive.CurrentUser)
        Dim rootNode = Me.CollectionsTreeView.Nodes(NameOf(Nodes.RootNode))
        For Each value In keyValues
            rootNode.Nodes.Add(New TreeNode With
                {
                    .Name = value,
                    .Text = value
                })

        Dim subKeyValues = registry.GetKeyValues(String.Format("{0}\{1}", My.Settings.RegistryPathCollections, value), RegistryHive.CurrentUser)
            For Each computerName In subKeyValues
                rootNode.Nodes(value).Nodes.Add(New TreeNode With
                        {
                            .Name = computerName,
                            .Text = computerName,
                            .ContextMenuStrip = collectionNodeMenuStrip
                        })
            Next
        Next

        Me.CollectionsTreeView.Nodes(NameOf(Nodes.RootNode)).Expand()
    End Sub

    Private Sub btnStart_Click(sender As Object, e As EventArgs) Handles btnStart.Click
        If Me.CollectionsTreeView.SelectedNode Is Nothing Then
            MsgBox("No collection selected")
            Exit Sub
        End If

        If Me.CollectionsTreeView.SelectedNode.Name = NameOf(Nodes.Collections) OrElse Me.CollectionsTreeView.SelectedNode.Parent.Name <> NameOf(Nodes.Collections) Then
            MsgBox("Invalid collection selected")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(Me.FileComboBox.Text) Then
            MsgBox("Invalid file selected")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(Me.TaskNameTextBox.Text) Then
            MsgBox("No task name selected")
            Exit Sub
        End If

        For Each q1ResultNode As TreeNode In OwnerForm.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Results)).Nodes
            If q1ResultNode.Text = Me.TaskNameTextBox.Text Then
                MsgBox(String.Format("Task: {0} already exists", q1ResultNode.Text))
                Exit Sub
            End If
        Next

        If My.Settings.RecentTasks IsNot Nothing Then
            Dim addFile As Boolean = True
            For Each i As String In My.Settings.RecentTasks
                If i = Me.FileComboBox.Text Then addFile = False
            Next

            If addFile Then
                My.Settings.RecentTasks.Add(Me.FileComboBox.Text)
                Me.FileComboBox.Items.Add(Me.FileComboBox.Text)
            End If

            If My.Settings.RecentTasks.Count > 10 Then
                Do Until My.Settings.RecentTasks.Count = 10
                    My.Settings.RecentTasks.RemoveAt(0)
                    Me.FileComboBox.Items.RemoveAt(0)
                Loop
            End If
        Else
            My.Settings.RecentTasks = New Specialized.StringCollection()
            My.Settings.RecentTasks.Add(Me.FileComboBox.Text)
            Me.FileComboBox.Items.Add(Me.FileComboBox.Text)
        End If

        My.Settings.Save()

        Dim computers As New List(Of String)
        For Each selectedTreeNode As TreeNode In CollectionsTreeView.SelectedNode.Nodes
            computers.Add(selectedTreeNode.Text)
        Next

        Dim taskResult As New TaskResultControl(computers, FileComboBox.Text, ArgumentsTextBox.Text, chbViewable.Checked, OwnerForm)
        taskResult.LabelName.Text = String.Format("Task: {0} on Collection: {1}", FileComboBox.Text, CollectionsTreeView.SelectedNode.Text)
        Dim q2ResultNode As New TreeNode With
            {
                .Text = Me.TaskNameTextBox.Text,
                .Tag = taskResult
            }

        Me.OwnerForm.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Results)).Nodes.Add(q2ResultNode)
        Me.OwnerForm.ResourceExplorer.SelectedNode = q2ResultNode
    End Sub

    Private Sub GetInfo()
        Dim searcher As New DataSourceSearcher(Me.CollectionsTreeView.SelectedNode.Text, Me.OwnerForm.BindingSource)

        Dim searchResults As List(Of Computer) = searcher.GetComputers()
        If searchResults IsNot Nothing Then
            Me.OwnerForm.UserInputComboBox.SelectedItem = searchResults(0)
            Me.OwnerForm.SubmitButton.PerformClick()
        End If
    End Sub

    Private Sub ArgumentsCheckBox_CheckedChanged(sender As Object, e As EventArgs) Handles ArgumentsCheckBox.CheckedChanged
        If Me.ArgumentsCheckBox.Checked Then
            Me.ArgumentsLabel.Visible = False
            Me.ArgumentsTextBox.Visible = True
        ElseIf Not Me.ArgumentsCheckBox.Checked Then
            Me.ArgumentsLabel.Visible = True
            Me.ArgumentsTextBox.Visible = False
            Me.ArgumentsTextBox.Text = String.Empty
        End If
    End Sub

    Private Sub BrowseButton_Click(sender As Object, e As EventArgs) Handles BrowseButton.Click
        Dim openDialog As New OpenFileDialog() With
        {
            .Title = "Select the file to be installed",
            .Filter = "All Files (*.*)|*.*|.bat files (*.bat)|*.bat|.exe files (*.exe)|*.exe|.msi files (*.msi)|*.msi|.msu files (*.msu)|*.msu|.cmd files (*.cmd)|*.cmd|printer export files (*.PrinterExport)|*.PrinterExport",
            .FilterIndex = 1,
            .Multiselect = False
        }

        If openDialog.ShowDialog() = DialogResult.Cancel Then
            Me.FileComboBox.SelectedItem = Me.FileComboBox.SelectedItem
            Exit Sub
        End If

        If Not openDialog.FileName.StartsWith("\\") Then
            Dim message = "System Tool has detected that this file is not from a shared location. Are you sure you want to continue?"
            If MsgBox(message, MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                Me.FileComboBox.Items.Add(openDialog.FileName)
                Me.FileComboBox.SelectedItem = openDialog.FileName
            Else
                Me.FileComboBox.SelectedItem = FileComboBox.SelectedItem
            End If
        Else
            Me.FileComboBox.Items.Add(openDialog.FileName)
            Me.FileComboBox.SelectedItem = openDialog.FileName
        End If
    End Sub

End Class
