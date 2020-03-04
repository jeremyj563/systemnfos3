Imports System.ComponentModel
Imports System.Reflection
Imports System.Text.RegularExpressions

Public Class UpdatesTab
    Inherits Tab

    Private Property UpdatesListView As ListView

    Public Sub New(ownerTab As TabControl, computerContext As ComputerControl)
        MyBase.New(ownerTab, computerContext)

        Me.Text = "Windows Updates"
        AddHandler LoaderBackgroundThread.DoWork, AddressOf UpdatePageEnumerator
        AddHandler ExportBackgroundThread.DoWork, AddressOf UpdateInfoExport
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgSU As String = "Security Updates"
        Public Shared ReadOnly Property lsvgUP As String = "Updates"
        Public Shared ReadOnly Property lsvgSP As String = "Service Packs"
        Public Shared ReadOnly Property lsvgHF As String = "Hotfixes"
        Public Shared ReadOnly Property lsvgOT As String = "Other"
    End Structure

    Private Sub UpdatePageEnumerator(sender As Object, e As DoWorkEventArgs)
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Me.UpdatesListView = NewBasicInfoListView(3)
        With Me.UpdatesListView.Groups
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgSU), ListViewGroups.lsvgSU))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgUP), ListViewGroups.lsvgUP))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgSP), ListViewGroups.lsvgSP))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgHF), ListViewGroups.lsvgHF))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgOT), ListViewGroups.lsvgOT))
        End With

        Dim updates As ManagementObjectCollection = Me.ComputerContext.WMI.Query("SELECT HotfixID, InstalledOn, Caption, Description FROM Win32_QuickFixEngineering")
        For Each update As ManagementBaseObject In updates
            If MyBase.UserCancellationPending() Then Exit Sub

            If update.Properties("HotfixID").Value.ToString().Trim().ToUpper() = "FILE 1" Then Continue For

            Dim groupName As String = Nothing
            Select Case update.Properties("Description").Value.ToString().Trim().ToUpper()
                Case "SECURITY UPDATE"
                    groupName = NameOf(ListViewGroups.lsvgSU)
                Case "UPDATE"
                    groupName = NameOf(ListViewGroups.lsvgUP)
                Case "SERVICE PACK"
                    groupName = NameOf(ListViewGroups.lsvgSP)
                Case "HOTFIX"
                    groupName = NameOf(ListViewGroups.lsvgHF)
                Case Else
                    groupName = NameOf(ListViewGroups.lsvgOT)
            End Select

            If update.Properties("InstalledOn").Value = "" Then
                NewTabWriterItem(update.Properties("HotfixID").Value, New Object() {"1/1/1601", update.Properties("Caption").Value, update}, groupName)
            ElseIf New Regex("^[a-zA-Z0-9]+$").IsMatch(update.Properties("HotFixID").Value) Then

                NewTabWriterItem(update.Properties("HotfixID").Value, New Object() {update.Properties("InstalledOn").Value, update.Properties("Caption").Value, update}, groupName)
            End If
        Next

        Me.UpdatesListView.Items.AddRange(NewBaseListViewItems(Me.UpdatesListView, Me.TabWriterObjects.ToArray()))
        Me.UpdatesListView.Sorting = SortOrder.Descending

        Me.UpdatesListView.ListViewItemSorter = New Comparer(1, True, Comparer.SortTypes.Descending)
        Me.UpdatesListView.Sort()

        Dim updatePanel As New Panel With
            {
                .Width = Me.Width,
                .Height = Me.Height - 25,
                .Anchor = AnchorStyles.Top + AnchorStyles.Left + AnchorStyles.Bottom + AnchorStyles.Right
            }

        ' Create a search textbox
        Dim searchTextBox As TextBox = NewSearchTextBox("Enter your update search here")

        ' Clear all controls from the tab
        Me.InvokeClearControls()

        ' Add the panel to the tab
        Me.InvokeAddControl(updatePanel)

        ' Display the Application listing in the first panel of the split container
        ShowListView(Me.UpdatesListView, updatePanel)

        ' Place the search textbox in the bottom of the tab
        searchTextBox.Location = New Point(0, Me.Height - 25)
        Me.InvokeAddControl(searchTextBox)

        AddHandler Me.UpdatesListView.ContextMenuStripChanged, AddressOf UpdateAddMenuStripHandler
    End Sub

    Private Sub UpdateAddMenuStripHandler(sender As Object, e As EventArgs)
        ' Creates additional menu options on right click
        If Me.MainListView.SelectedItems.Count > 0 Then
            Dim update As ManagementBaseObject = Me.MainListView.SelectedItems(0).Tag
            If update.Properties("Caption").Value IsNot Nothing OrElse Not String.IsNullOrWhiteSpace(update.Properties("Caption").Value) Then
                Dim getInfoMenuItem As New ToolStripMenuItem("Get Info", Nothing, AddressOf DisplaySite)
                getInfoMenuItem.Tag = GetSelectedListViewItem(MainListView).Tag.Properties("Caption").Value
                AddToCurrentMenuStrip(getInfoMenuItem)
            Else
                AddToCurrentMenuStrip(New ToolStripMenuItem("No Information Available"))
            End If

            AddToCurrentMenuStrip(New ToolStripSeparator)
        End If
    End Sub

    Private Sub DisplaySite(sender As Object, e As EventArgs)
        Try
            If TypeOf sender Is ToolStripMenuItem Then
                Dim menuItem As ToolStripMenuItem = sender

                If TypeOf menuItem.Tag Is String Then
                    Dim path As String = menuItem.Tag
                    Process.Start(path)
                End If
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try
    End Sub

    Private Sub UpdateInfoExport(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.UpdatesListView, userSelectedCSVFilePath)
    End Sub

End Class