Imports Microsoft.Win32
Imports System.IO
Imports System.Reflection
Imports System.ComponentModel

Public Class CollectionOptionsControl

    Public Property OwnerForm As MainForm

    Private Property RegistryContext As New RegistryController
    Private Property ComputerNamesToAdd As List(Of String)

    Private WithEvents InitWorker As New BackgroundWorker

    Public Sub New(ownerForm As MainForm)
        InitializeComponent()
        Me.OwnerForm = ownerForm
    End Sub

    Public Sub AddToCollectionListView(computerNames As List(Of String))
        Dim currentlyLoadedComputers As New List(Of String)
        For Each listViewItem As ListViewItem In ComputerStatusListView.Items
            currentlyLoadedComputers.Add(listViewItem.Text.Trim().ToUpper())
        Next

        For Each computerName As String In computerNames
            If Not currentlyLoadedComputers.Contains(computerName.Trim().ToUpper()) Then
                Dim listViewItem As New ListViewItem With {.Text = computerName}
                Me.ComputerStatusListView.Items.Add(listViewItem)
            End If
        Next
    End Sub

    Private Sub CollectionOptionsControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim computerStatusMenuStrip As New ContextMenuStrip
        Dim activeForm As MainForm = Form.ActiveForm

        Dim collectionRegistryPath As String = Path.Combine(My.Settings.RegistryPathCollections, activeForm.ResourceExplorer.SelectedNode.Text)
        Dim computerNames As List(Of String) = Me.RegistryContext.GetKeyValues(collectionRegistryPath, RegistryHive.CurrentUser).ToList()
        AddToCollectionListView(computerNames)

        With computerStatusMenuStrip.Items
            .Add("Get Info", Nothing, AddressOf GetInfo)
            .Add("Remove", Nothing, AddressOf RemoveFromListViewSystemStatus)
        End With

        With Me.ComputerStatusListView
            .ContextMenuStrip = computerStatusMenuStrip
            .Columns(0).Width = 171
        End With
    End Sub

    Private Sub RemoveFromListViewSystemStatus(sender As Object, e As EventArgs)
        If Me.ComputerStatusListView.SelectedItems.Count > 0 Then
            For Each listViewItem As ListViewItem In Me.ComputerStatusListView.SelectedItems
                Dim collectionRegistryPath As String = Path.Combine(My.Settings.RegistryPathCollections, Me.ComputerStatusListView.Columns(0).Text)
                Dim subKeyToRemovePath = Path.Combine(collectionRegistryPath, listViewItem.Text)

                My.Computer.Registry.CurrentUser.DeleteSubKey(subKeyToRemovePath)
                listViewItem.Remove()
            Next
        End If
    End Sub

    Private Sub ButtonSelectList_Click(sender As Object, e As EventArgs) Handles ButtonSelectList.Click
        Dim openDialog As New OpenFileDialog With
            {
                .InitialDirectory = SystemDrive,
                .Title = "Select text file that lists the computers to be added to the collection",
                .Filter = "Text Files (*.txt)|*.txt",
                .Multiselect = False
            }

        Dim dialogResult = openDialog.ShowDialog()
        If dialogResult = DialogResult.Cancel Then
            Exit Sub
        End If

        For Each listViewItem As ListViewItem In Me.ComputerNamesListView.Items
            If listViewItem IsNot Nothing Then
                Me.ComputerNamesListView.Items.Remove(listViewItem)
            End If
        Next

        Using reader As New StreamReader(openDialog.FileName)
            Dim line As String = reader.ReadLine()
            Dim columnNumber As Integer = 0
            If line.Contains(ControlChars.Tab) Then
                Dim columnNames As String() = line.Split(ControlChars.Tab)
                If columnNames.Length > 1 Then
                    Dim userSelectedColumnNumber As Integer = InputBox(String.Format("This list contains multiple columns.{0}Which column would you like to use?", Environment.NewLine), "Select column:", "1")
                    If userSelectedColumnNumber <= columnNames.Length Then
                        columnNumber = userSelectedColumnNumber
                    End If
                End If
            End If

            Do While line IsNot Nothing
                If columnNumber <= 0 Then
                    Me.ComputerNamesListView.Items.Add(line.Split(ControlChars.Tab)(columnNumber).Trim())
                Else
                    Me.ComputerNamesListView.Items.Add(line.Split(ControlChars.Tab)(columnNumber - 1).Trim())
                End If

                line = reader.ReadLine()
            Loop

            columnNumber = Nothing
        End Using

        For Each listViewItem As ListViewItem In Me.ComputerNamesListView.Items
            listViewItem.Checked = True
        Next
    End Sub

    Private Sub AddButton_Click(sender As Object, e As EventArgs) Handles AddButton.Click
        If String.IsNullOrWhiteSpace(Me.SystemTextBox.Text) Then
            MsgBox("You must enter a system name or IP address")
        Else
            Me.StatusTwoLabel.Text = String.Empty

            For Each listViewItem As ListViewItem In Me.ComputerStatusListView.Items
                If listViewItem.Text = Me.SystemTextBox.Text Then
                    MsgBox(String.Format("{0} is already a member of collection: {1}", Me.SystemTextBox.Text, Me.ComputerStatusListView.Columns(0).Text))
                    Me.StatusTwoLabel.Text = String.Format("{0} not added to the collection", Me.SystemTextBox.Text)
                    Me.SystemTextBox.Text = String.Empty
                    Me.SystemTextBox.Focus()
                    Exit Sub
                End If
            Next

            Try
                Dim collectionRegistryPath As String = Path.Combine(My.Settings.RegistryPathCollections, Me.ComputerStatusListView.Columns(0).Text)
                Dim subKeyToCreatePath As String = Path.Combine(collectionRegistryPath, Me.SystemTextBox.Text)
                Me.RegistryContext.NewKey(subKeyToCreatePath, RegistryHive.CurrentUser)

                Dim listViewItem As New ListViewItem() With {.Text = Me.SystemTextBox.Text}
                Me.ComputerStatusListView.Items.Add(listViewItem)
                Me.StatusTwoLabel.Text = String.Format("{0} added to Collection", Me.SystemTextBox.Text)
                Me.SystemTextBox.Text = String.Empty
                Me.SystemTextBox.Focus()
            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                MsgBox(ex.Message, icon:=MessageBoxIcon.Error)
                Me.StatusTwoLabel.Text = String.Format("{0} not added to collection", Me.SystemTextBox.Text)
            End Try
        End If
    End Sub

    Private Sub ImportButton_Click(sender As Object, e As EventArgs) Handles ImportButton.Click
        Me.ComputerNamesToAdd = New List(Of String)()
        If Me.ComputerNamesListView.CheckedItems.Count = 0 Then
            MsgBox("No systems selected")
        Else
            For Each listViewItem As ListViewItem In Me.ComputerNamesListView.CheckedItems
                Me.ComputerNamesToAdd.Add(listViewItem.Text)
            Next

            Me.StatusLabel.Text = "Importing..."

            With Me.InitWorker
                .WorkerReportsProgress = True
                .WorkerSupportsCancellation = True
                .RunWorkerAsync()
            End With
        End If
    End Sub

    Private Sub InitWorker_DoWork(sender As Object, e As DoWorkEventArgs) Handles InitWorker.DoWork
        Try
            For Each computerName As String In Me.ComputerNamesToAdd
                computerName = computerName.Replace(ControlChars.Tab, String.Empty).Trim()
                Try
                    Dim collectionRegistryPath = Path.Combine(My.Settings.RegistryPathCollections, ComputerStatusListView.Columns(0).Text)
                    Dim subKeyToCreatePath = Path.Combine(collectionRegistryPath, computerName)
                    Me.RegistryContext.NewKey(subKeyToCreatePath, RegistryHive.CurrentUser)
                Catch ex As Exception
                    LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                    Dim message = String.Format("Import failed for system: {0} Below is the error message: {1}{2}", Me.ComputerName, Environment.NewLine, ex.Message)
                    MsgBox(message, icon:=MessageBoxIcon.Error)
                End Try
            Next
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub InitWorker_RunWorkerCompleted(sender As Object, e As RunWorkerCompletedEventArgs) Handles InitWorker.RunWorkerCompleted
        AddToCollectionListView(Me.ComputerNamesToAdd)
        Me.ComputerNamesToAdd.Clear()
        Me.StatusLabel.Text = "Import complete"
    End Sub

    Private Sub ComputerNamesListView_Resize(sender As Object, e As EventArgs) Handles ComputerNamesListView.Resize
        Me.ComputerNamesListView.Columns(0).Width = Me.ComputerNamesListView.Width - 5
    End Sub

    Private Sub GetInfo()
        For Each listViewItem As ListViewItem In ComputerStatusListView.SelectedItems
            Dim searcher As New DataSourceSearcher(listViewItem.Text, Me.OwnerForm.BindingSource)
            Dim searchResults As List(Of Computer) = searcher.GetComputers()
            If searchResults IsNot Nothing Then
                Me.OwnerForm.UserInputComboBox.SelectedItem = searchResults.First()
                Me.OwnerForm.SubmitButton.PerformClick()
            End If
        Next
    End Sub

End Class