Imports Microsoft.Win32

Public Class QueryControl

    Public Property OwnerForm As MainForm

    Public Sub New(ownerForm As MainForm)
        InitializeComponent()
        Me.OwnerForm = ownerForm
    End Sub

    Private Sub QueryControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        LoadCollectionInfo()
        Me.QueryListView.Columns(0).Width = 214
        Me.QueryListView.Columns(1).Width = 67

        Dim queryMenuStrip As New ContextMenuStrip()
        queryMenuStrip.Items.Add("Remove", Nothing, Sub()
                                                        For Each listViewItem As ListViewItem In Me.QueryListView.SelectedItems
                                                            listViewItem.Remove()
                                                        Next
                                                    End Sub)

        Me.QueryListView.ContextMenuStrip = queryMenuStrip
        Me.WMIRadioButton.Checked = True
    End Sub

    Public Sub LoadCollectionInfo()
        Me.CollectionsTreeView.Nodes.Clear()

        Dim collectionsMenuStrip As New ContextMenuStrip()
        collectionsMenuStrip.Items.Add("Get Info", Nothing, AddressOf GetInfo)

        Me.CollectionsTreeView.Nodes.Add(New TreeNode With
        {
            .Name = NameOf(Nodes.RootNode),
            .Text = NameOf(Nodes.Collections)
        })

        Dim registry As New RegistryController()
        Dim collectionsRootNode = Me.CollectionsTreeView.Nodes(NameOf(Nodes.RootNode))
        For Each value In registry.GetKeyValues(My.Settings.RegistryPathCollections, RegistryHive.CurrentUser)
            collectionsRootNode.Nodes.Add(New TreeNode With
            {
                .Name = value,
                .Text = value
            })

            For Each computerName In registry.GetKeyValues($"{My.Settings.RegistryPathCollections}\{value}", RegistryHive.CurrentUser)
                collectionsRootNode.Nodes(value).Nodes.Add(New TreeNode With
                {
                    .Name = computerName,
                    .Text = computerName,
                    .ContextMenuStrip = collectionsMenuStrip
                })
            Next
        Next

        collectionsRootNode.Expand()
    End Sub

    Private Sub WMIRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles WMIRadioButton.CheckedChanged
        If Me.WMIRadioButton.Checked Then
            With Me.AttributeClassComboBox
                .Items.Clear()
                .Items.Add("Network Adapter Configuration")
                .Items.Add("Operating System")
                .Items.Add("System Information")
                .Items.Add("Dell Warranty Information")
            End With
        End If
    End Sub

    Private Sub LDAPRadioButton_CheckedChanged(sender As Object, e As EventArgs) Handles LDAPRadioButton.CheckedChanged
        If Me.LDAPRadioButton.Checked Then
            With Me.AttributeClassComboBox
                .Items.Clear()
                .Items.Add("Network Adapter Configuration")
                .Items.Add("Operating System")
                .Items.Add("System Information")
            End With
        End If
    End Sub

    Private Sub AttributeClassComboBox_SelectedIndexChanged(sender As Object, e As EventArgs) Handles AttributeClassComboBox.SelectedIndexChanged
        Me.AttributeComboBox.Items.Clear()

        With Me.AttributeComboBox
            If Me.WMIRadioButton.Checked Then
                Select Case Me.AttributeClassComboBox.SelectedItem

                    Case "Network Adapter Configuration"
                        .Items.Add("NAC IP Address")
                        .Items.Add("NAC MAC Address")
                        .Items.Add("NAC Default Gateway")

                    Case "Operating System"
                        .Items.Add("OS Name")
                        .Items.Add("OS Description")
                        .Items.Add("OS Architecture")

                    Case "System Information"
                        .Items.Add("CS Model")
                        .Items.Add("CS Serial Number")
                        .Items.Add("CS Bios Version")
                        .Items.Add("CS Image Date")
                        .Items.Add("CS Bitlocker Status")
                        .Items.Add("CS Last Boot")
                        .Items.Add("CS Total HD Space")
                        .Items.Add("CS Free HD Space")
                        .Items.Add("CS CPU")
                        .Items.Add("CS Total Memory")
                        .Items.Add("CS HDD Serial")
                End Select

            ElseIf Me.LDAPRadioButton.Checked Then
                Select Case Me.AttributeClassComboBox.SelectedItem

                    Case "Network Adapter Configuration"
                        .Items.Add("NAC IP Address")
                        .Items.Add("NAC MAC Address")

                    Case "Operating System"
                        .Items.Add("OS Name")
                        .Items.Add("OS Description")
                        .Items.Add("OS Architecture")

                    Case "System Information"
                        .Items.Add("CS Location")
                        .Items.Add("CS Model")
                        .Items.Add("CS Serial Number")
                        .Items.Add("CS Bios Version")
                        .Items.Add("CS Image Date")
                        .Items.Add("CS Location")
                        .Items.Add("CS Last Logged On User")
                End Select
            End If
        End With
    End Sub

    Private Sub AddButton_Click(sender As Object, e As EventArgs) Handles AddButton.Click
        Dim itemToAdd As String = Nothing
        If Me.WMIRadioButton.Checked Then itemToAdd = "WMI"
        If Me.LDAPRadioButton.Checked Then itemToAdd = "LDAP"

        For Each listViewItem As ListViewItem In Me.QueryListView.Items
            If listViewItem.Text = AttributeComboBox.SelectedItem And listViewItem.SubItems(1).Text = itemToAdd Then
                MsgBox("This attribute is already selected!")
                Exit Sub
            End If
        Next

        If AttributeClassComboBox.SelectedItem = Nothing Then
            MsgBox("You must select a valid attribute class!")
        ElseIf AttributeComboBox.SelectedItem = Nothing Then
            MsgBox("You must select a valid attribute!")
        Else
            Dim listViewItem As New ListViewItem
            listViewItem.Text = AttributeComboBox.SelectedItem
            listViewItem.SubItems.Add(itemToAdd)

            Me.QueryListView.Items.Add(listViewItem)
        End If
    End Sub

    Private Sub StartButton_Click(sender As Object, e As EventArgs) Handles StartButton.Click
        If Me.CollectionsTreeView.SelectedNode Is Nothing Then
            MsgBox("No collection selected")
            Exit Sub
        End If

        If Me.CollectionsTreeView.SelectedNode.Name = NameOf(Nodes.RootNode) OrElse Me.CollectionsTreeView.SelectedNode.Parent.Name <> NameOf(Nodes.RootNode) Then
            MsgBox("Invalid collection selected")
            Exit Sub
        End If

        If Me.QueryListView.Items.Count = 0 Then
            MsgBox("No query items selected")
            Exit Sub
        End If

        If String.IsNullOrWhiteSpace(Me.QueryNameTextBox.Text) Then
            MsgBox("No query name selected")
            Exit Sub
        End If

        For Each query1TreeNode As TreeNode In Me.OwnerForm.ResourceExplorer.Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Results)).Nodes
            If query1TreeNode.Text = Me.QueryNameTextBox.Text Then
                MsgBox($"Query: {query1TreeNode.Text} already exists")
                Exit Sub
            End If
        Next

        Dim computers As New List(Of String)
        For Each selectedTreeNode As TreeNode In Me.CollectionsTreeView.SelectedNode.Nodes
            computers.Add(selectedTreeNode.Text)
        Next

        Dim properties As New List(Of String)
        For Each [property] As ListViewItem In QueryListView.Items
            properties.Add($"{[property].Text}>>{[property].SubItems(1).Text}")
        Next

        Dim query2Result As New QueryResultControl(computers, properties, Me.OwnerForm)
        query2Result.lblName.Text = $"Query: {Me.QueryNameTextBox.Text} for Collection: {Me.CollectionsTreeView.SelectedNode.Text}"
        Dim query2TreeNode As New TreeNode With
            {
                .Text = Me.QueryNameTextBox.Text,
                .Tag = query2Result
            }

        With Me.OwnerForm.ResourceExplorer
            .Nodes(NameOf(Nodes.RootNode)).Nodes(NameOf(Nodes.Results)).Nodes.Add(query2TreeNode)
            .SelectedNode = query2TreeNode
        End With
    End Sub

    Private Sub GetInfo()
        Dim searcher As New DataSourceSearcher(Me.CollectionsTreeView.SelectedNode.Text, Me.OwnerForm.BindingSource)
        Dim searchResults As List(Of Computer) = searcher.GetComputers()
        If searchResults IsNot Nothing Then
            Me.OwnerForm.UserInputComboBox.SelectedItem = searchResults.First()
            Me.OwnerForm.SubmitButton.PerformClick()
        End If
    End Sub

End Class
