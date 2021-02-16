Imports System.ComponentModel
Imports System.IO
Imports System.Reflection

Public Class ProfileTab
    Inherits BaseTab

    Private Property ProfileInfoListView As ListView

    Public Sub New(ownerTab As TabControl, computerPanel As ComputerPanel)
        MyBase.New(ownerTab, computerPanel)

        Me.Text = "Profiles"
        AddHandler InitWorker.DoWork, AddressOf InitializeProfileTab
        AddHandler ExportWorker.DoWork, AddressOf ExportProfileInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgUS As String = "User Profiles / Last Access Time"
    End Structure

    Private Sub InitializeProfileTab()
        ' Clear the tab of all child controls
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Me.ProfileInfoListView = NewBaseListView(3)

        ' Create the ListView Groups
        Me.ProfileInfoListView.Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgUS), ListViewGroups.lsvgUS))

        For Each user In Directory.GetDirectories($"\\{Me.ComputerPanel.Computer.ConnectionString}\C$\Users")
            If MyBase.UserCancellationPending() Then Exit Sub

            AddTabWriterItem(New DirectoryInfo(user).Name, New String() {New DirectoryInfo(user).LastAccessTime, user}, NameOf(ListViewGroups.lsvgUS))
        Next

        ' Add Items into ListView
        Me.ProfileInfoListView.Items.AddRange(NewBaseListViewItems(Me.ProfileInfoListView, TabWriterObjects.ToArray()))
        ShowListView(Me.ProfileInfoListView, Me)

        ' Create additional ContextMenuStripItems
        AddHandler Me.ProfileInfoListView.ContextMenuStripChanged, AddressOf OnProfileMenuStripChanged
    End Sub

    Private Sub OnProfileMenuStripChanged(sender As Object, e As EventArgs)
        Dim sortMenuItem As New ToolStripMenuItem("Sort")

        If Me.CurrentListView.SelectedItems.Count > 0 Then
            Dim openProfileMenuItem As New ToolStripMenuItem("Open Profile", Nothing, AddressOf OpenProfile) With {
                .Tag = Me.CurrentListView.SelectedItems(0).SubItems(2).Text
            }

            With sortMenuItem.DropDownItems
                .Add("Name", Nothing, AddressOf OrderByName)
                .Add("Date", Nothing, AddressOf OrderByDate)
            End With

            AddToCurrentMenuStrip(openProfileMenuItem)
            AddToCurrentMenuStrip(New ToolStripSeparator)
            AddToCurrentMenuStrip(sortMenuItem)
            AddToCurrentMenuStrip(New ToolStripSeparator)
        Else
            With sortMenuItem.DropDownItems
                .Add("Name", Nothing, AddressOf OrderByName)
                .Add("Date", Nothing, AddressOf OrderByDate)
            End With

            AddToCurrentMenuStrip(sortMenuItem)
            AddToCurrentMenuStrip(New ToolStripSeparator)
        End If

    End Sub

    Private Sub OrderByName()
        With Me.CurrentListView
            .ListViewItemSorter = New Comparer(0, False, Comparer.SortTypes.Ascending)
            .Sort()
        End With
    End Sub

    Private Sub OrderByDate()
        With Me.CurrentListView
            .ListViewItemSorter = New Comparer(1, True, Comparer.SortTypes.Descending)
            .Sort()
        End With
    End Sub

    Private Sub OpenProfile(sender As Object, e As EventArgs)
        Try
            If Directory.Exists(sender.tag) Then
                Process.Start(sender.tag)
            Else
                Dim message = $"Unable to locate the requested location: {sender.tag}"
                Me.ComputerPanel.WriteMessage(message, Color.Red)
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub ExportProfileInfo(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.ProfileInfoListView, userSelectedCSVFilePath)
    End Sub

End Class
