Imports System.ComponentModel
Imports System.IO
Imports System.Reflection

Public Class ProfileTab
    Inherits Tab

    Private Property ProfileInfoListView As ListView

    Public Sub New(ownerTab As TabControl, computerContext As ComputerControl)
        MyBase.New(ownerTab, computerContext)

        Me.Text = "Profiles"
        AddHandler LoaderBackgroundThread.DoWork, AddressOf InitializeProfileTab
        AddHandler ExportBackgroundThread.DoWork, AddressOf ExportProfileInfo
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

        Me.ProfileInfoListView = NewBasicInfoListView(3)

        ' Create the ListView Groups
        Me.ProfileInfoListView.Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgUS), ListViewGroups.lsvgUS))

        For Each user As String In Directory.GetDirectories(String.Format("\\{0}\C$\Users", Me.ComputerContext.Computer.ConnectionString))
            If MyBase.UserCancellationPending() Then Exit Sub

            NewTabWriterItem(New DirectoryInfo(user).Name, New String() {New DirectoryInfo(user).LastAccessTime, user}, NameOf(ListViewGroups.lsvgUS))
        Next

        ' Add Items into ListView
        Me.ProfileInfoListView.Items.AddRange(NewBaseListViewItems(Me.ProfileInfoListView, TabWriterObjects.ToArray()))
        ShowListView(Me.ProfileInfoListView, Me)

        ' Create additional ContextMenuStripItems
        AddHandler Me.ProfileInfoListView.ContextMenuStripChanged, AddressOf OnProfileMenuStripChanged
    End Sub

    Private Sub OnProfileMenuStripChanged(sender As Object, e As EventArgs)
        Dim sortMenuItem As New ToolStripMenuItem("Sort")

        If Me.MainListView.SelectedItems.Count > 0 Then
            Dim openProfileMenuItem As New ToolStripMenuItem("Open Profile", Nothing, AddressOf OpenProfile)
            openProfileMenuItem.Tag = Me.MainListView.SelectedItems(0).SubItems(2).Text

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
        With Me.MainListView
            .ListViewItemSorter = New Comparer(0, False, Comparer.SortTypes.Ascending)
            .Sort()
        End With
    End Sub

    Private Sub OrderByDate()
        With Me.MainListView
            .ListViewItemSorter = New Comparer(1, True, Comparer.SortTypes.Descending)
            .Sort()
        End With
    End Sub

    Private Sub OpenProfile(sender As Object, e As EventArgs)
        Try
            If Directory.Exists(sender.tag) Then
                Process.Start(sender.tag)
            Else
                Dim message = String.Format("Unable to locate the requested location: {0}", sender.tag)
                Me.ComputerContext.WriteMessage(message, Color.Red)
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub ExportProfileInfo(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.ProfileInfoListView, userSelectedCSVFilePath)
    End Sub

End Class
