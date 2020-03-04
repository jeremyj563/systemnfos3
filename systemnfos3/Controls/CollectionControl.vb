Imports System.ComponentModel
Imports System.IO
Imports Microsoft.Win32

Public Class CollectionControl

    Public Property OwnerForm As MainForm
    Public Property OwnerNode As TreeNode

    Private Property CollectionListView As ListView
    Private Property BackgroundThread As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
    Private Property RegistryKey As String
    Private Property BoundData As BindingSource
    Private Property SavedInRegsitry As Boolean = False
    Private Property Searcher As DataSourceSearcher

#Region " Constructors "
    Public Sub New(ownerForm As MainForm)
        ' This call is required by the designer.
        InitializeComponent()

        Me.OwnerForm = ownerForm

        Dim collectionListView = NewCollectionListView()
        Me.InvokeAddControl(collectionListView)
    End Sub

    Public Sub New(ownerForm As MainForm, ownerNode As TreeNode, searcher As DataSourceSearcher)
        ' Constructor for new collections made from user's search query

        ' This call is required by the designer.
        InitializeComponent()

        Me.OwnerForm = ownerForm
        Me.OwnerNode = ownerNode
        Me.Searcher = searcher

        AddHandler Me.BackgroundThread.DoWork, AddressOf EnumerateSearcherComputers
    End Sub

    Public Sub New(ownerForm As MainForm, collectionRegistryKey As String, boundData As BindingSource)
        ' Constructor for existing collections stored in the Registry

        ' This call is required by the designer.
        InitializeComponent()

        Me.OwnerForm = ownerForm
        Me.RegistryKey = collectionRegistryKey
        Me.BoundData = boundData
        Me.SavedInRegsitry = True

        AddHandler Me.BackgroundThread.DoWork, AddressOf EnumerateRegistryComputers
    End Sub

#End Region

#Region " Collection Enumeration "

    Public Sub BeginEnumeration()
        If Me.BackgroundThread.IsBusy Then
            Me.BackgroundThread.CancelAsync()
        End If

        Me.BackgroundThread.RunWorkerAsync()
    End Sub

    Private Sub EnumerateSearcherComputers()
        ' This is a new collection so create a Searcher instance and begin looking for computers

        Dim collectionListView As ListView = NewCollectionListView()

        Dim searchResults As List(Of Computer) = Me.Searcher.GetComputers()
        If searchResults IsNot Nothing And searchResults.Count > 1 Then

            For Each computer As Computer In searchResults
                Dim item As New ListViewItem(computer.Value)
                With item
                    .SubItems.Add(computer.Description)
                    .SubItems.Add(computer.UserName)
                    .SubItems.Add(computer.DisplayName)
                    .SubItems.Add(computer.LastLogon)
                    .SubItems.Add(computer.IPAddress)
                    .Tag = computer
                End With

                collectionListView.Items.Add(item)
            Next
        End If

        Me.UIThread(Sub() collectionListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent))
        Me.InvokeAddControl(collectionListView)
    End Sub

    Private Sub EnumerateRegistryComputers()
        ' This is an existing collection so load the computer names from the registry and then verify them against the verification data source

        Dim collectionListView As ListView = NewCollectionListView()

        Dim collectionRegistryKeyPath As String = Path.Combine(My.Settings.RegistryPathCollections, RegistryKey)
        Dim computerNames As String() = New RegistryController().GetKeyValues(collectionRegistryKeyPath, RegistryHive.CurrentUser)

        For Each computerName In computerNames
            Me.Searcher = New DataSourceSearcher(computerName, Me.BoundData)
            Dim computer As Computer = Searcher.GetComputer()
            If computer IsNot Nothing Then
                Dim listViewItem As New ListViewItem(computer.Value)
                With listViewItem
                    .SubItems.Add(computer.Description)
                    .SubItems.Add(computer.UserName)
                    .SubItems.Add(computer.DisplayName)
                    .SubItems.Add(computer.LastLogon)
                    .SubItems.Add(computer.IPAddress)
                    .Tag = computer
                End With

                collectionListView.Items.Add(listViewItem)
            Else
                Dim listViewItem As New ListViewItem(computerName)
                With listViewItem
                    .SubItems.Add("This object does not exist in LDAP")
                    .Tag = Nothing
                End With

                collectionListView.Items.Add(listViewItem)
            End If
        Next

        Me.UIThread(Sub() collectionListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent))
        Me.InvokeAddControl(collectionListView)
    End Sub

#End Region

#Region " Context Menu Handlers "

    Private Sub DisplayCollectionControlMenuStrip(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Right Then
            Dim listView As ListView = sender
            Dim menuStrip As New ContextMenuStrip

            Select Case Me.SavedInRegsitry

                Case False
                    If listView.SelectedItems.Count = 1 AndAlso listView.SelectedItems(0).Tag IsNot Nothing Then
                        Dim getInfoMenuItem As New ToolStripMenuItem("Get Info", Nothing, AddressOf GetComputerInfo)
                        Dim refreshMenuItem As New ToolStripMenuItem("Refresh", Nothing, AddressOf RefreshCollection)
                        getInfoMenuItem.Tag = listView.SelectedItems
                        menuStrip.Items.Add(getInfoMenuItem)
                        menuStrip.Items.Add(refreshMenuItem)
                    End If
                    Dim saveCollectionAs As New ToolStripMenuItem("Save Collection As...", Nothing, AddressOf SaveCollection)
                    saveCollectionAs.Tag = listView

                Case True
                    If listView.SelectedItems.Count = 1 AndAlso listView.SelectedItems(0).Tag IsNot Nothing Then
                        Dim getInfoMenuItem As New ToolStripMenuItem("Get Info", Nothing, AddressOf GetComputerInfo)
                        getInfoMenuItem.Tag = listView.SelectedItems
                        With menuStrip
                            .Items.Add(getInfoMenuItem)
                            .Items.Add(New ToolStripSeparator)
                            .Items.Add("Delete Computer", Nothing, AddressOf DeleteComputer)
                            .Items.Add(New ToolStripSeparator)
                        End With
                    End If
                    With menuStrip
                        .Items.Add("Add Computer", Nothing, AddressOf AddComputer)
                        .Items.Add("Import List", Nothing, AddressOf ImportList)
                    End With
            End Select

            listView.ContextMenuStrip = menuStrip
        End If
    End Sub

    Public Function NewCollectionMenuStrip() As ContextMenuStrip
        Dim menustrip As New ContextMenuStrip()
        menustrip.Items.Add("Remove From List", Nothing, AddressOf RemoveOwnerNode)

        Return menustrip
    End Function

#End Region

#Region " Collection ListView "

    Private Function NewCollectionListView() As ListView
        Me.CollectionListView = New ListView()

        With Me.CollectionListView
            .Height = Me.Height
            .Width = Me.Width
            .Dock = DockStyle.Fill
            .Anchor = AnchorStyles.Bottom + AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Right
            .Columns.Add("Computer Name")
            .Columns.Add("Description")
            .Columns.Add("Username")
            .Columns.Add("Display Name")
            .Columns.Add("Last Logon")
            .Columns.Add("IP Address")
            .View = View.Details
            .MultiSelect = False
            .FullRowSelect = True
            .HeaderStyle = ColumnHeaderStyle.Nonclickable
        End With

        AddHandler Me.CollectionListView.MouseUp, AddressOf DisplayCollectionControlMenuStrip
        AddHandler Me.CollectionListView.MouseDoubleClick, AddressOf GetComputerInfo

        Return Me.CollectionListView
    End Function

    Private Sub AddComputer()
        ' To Do
    End Sub

    Private Sub DeleteComputer()
        ' To Do
    End Sub

    Private Sub SaveCollection(sender As Object, e As EventArgs)
        MsgBox(sender.Tag.GetType().Name)
    End Sub

    Private Sub ImportList()
        ' To Do
    End Sub

    Private Sub GetComputerInfo(sender As Object, e As EventArgs)
        Dim selectedItems = Nothing

        If TypeOf sender Is ToolStripMenuItem AndAlso sender.Tag IsNot Nothing Then
            selectedItems = sender.Tag
        End If

        If TypeOf sender Is ListView AndAlso sender.SelectedItems IsNot Nothing Then
            selectedItems = sender.SelectedItems
        End If

        If selectedItems IsNot Nothing Then
            For Each item As ListViewItem In selectedItems
                With Me.OwnerForm
                    .UserInputComboBox.SelectedItem = item.Tag
                    .SubmitButton.PerformClick()
                End With
            Next
        End If
    End Sub

    Private Sub RefreshCollection()
        Me.InvokeClearControls()
        BeginEnumeration()
    End Sub

#End Region

#Region " Misc "

    Private Sub RemoveOwnerNode()
        Me.OwnerNode.Remove()
    End Sub

#End Region

End Class
