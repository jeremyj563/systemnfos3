Imports System.Reflection
Imports System.ComponentModel
Imports System.Windows.Forms.ListViewItem

Public MustInherit Class Tab
    Inherits TabPage

    Public Property OwnerTab As TabControl
    Public Property ComputerContext As ComputerControl
    Public Property IsDisplayValidationNeeded As Boolean = True
    Public Property TabWriterObjects As List(Of Object)
    Public Property EnumWriterObjects As List(Of Object)

    Private Property IsMainListViewCreated As Boolean = False

    Friend Property LoadingProgressBar As ProgressBar
    Friend Property ExportingProgressBar As ProgressBar
    Friend Property MainListView As ListView
    Friend Property LastSelectedListViewItem As ListViewItem = Nothing
    Friend Property SearchTextBox As TextBox

    Friend LoaderBackgroundThread As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
    Friend WithEvents ExportBackgroundThread As New BackgroundWorker() With {.WorkerSupportsCancellation = True, .WorkerReportsProgress = True}
    Friend WithEvents SelectionBackgroundThread As New BackgroundWorker()

    Public Sub New(ownerTab As TabControl, computerContext As ComputerControl)
        ' Setting the ComputerContext to its sender is required before control initialization
        ' That is because the VisibleChanged events fire off before the control is loaded, causing validation errors.
        Me.ComputerContext = computerContext

        ' This call is required by the designer.
        InitializeComponent()

        Me.OwnerTab = ownerTab
        Me.BackColor = Color.White
    End Sub

    Public Function NewBasicInfoListView(columnsCount As Integer) As ListView
        Dim basicInfo As New ListView() With
            {
                .View = View.Details,
                .Anchor = AnchorStyles.Left + AnchorStyles.Right + AnchorStyles.Bottom + AnchorStyles.Top,
                .Dock = DockStyle.Fill,
                .MultiSelect = False,
                .HideSelection = False,
                .FullRowSelect = True,
                .HeaderStyle = ColumnHeaderStyle.None
            }

        For i = 1 To columnsCount
            basicInfo.Columns.Add(i)
        Next

        basicInfo.ContextMenuStrip = New ContextMenuStrip()
        If Not Me.IsMainListViewCreated Then
            Me.MainListView = basicInfo
            Me.IsMainListViewCreated = True
        End If

        AddHandler basicInfo.MouseUp, AddressOf NewBaseMenuStrip

        Return basicInfo
    End Function

    Public Sub NewBaseMenuStrip(sender As Object, e As MouseEventArgs)
        If e.Button = Windows.Forms.MouseButtons.Right Then
            Dim listView As ListView = sender
            Dim menuStrip As New ContextMenuStrip()

            ' Add context menu items only available when an item is selected
            If listView.SelectedItems.Count = 1 Then
                ' Build "Copy" toolstrip menu item which allows copying the listview items to the clipboard
                Dim selectedItem As ListViewItem = listView.SelectedItems(0)
                Dim selectedItemText As String = Nothing

                ' Loop through all the available listview sub items and create a toolstrip dropdown item for each
                Dim copyMenuItem As New ToolStripMenuItem("Copy")
                For Each subItem As ListViewSubItem In selectedItem.SubItems
                    If String.IsNullOrWhiteSpace(subItem.Text) Then Continue For

                    ' Concatenate all items into a full string which will be added to the bottom of the dropdown items list
                    Dim subItemText As String = subItem.Text
                    selectedItemText += subItemText + " "

                    ' Add the toolstrip dropdown item for this listview sub item
                    copyMenuItem.DropDownItems.Add(subItemText.Trim(), Nothing, AddressOf CopyToolStripItemTextToClipboard)
                Next

                If Not String.IsNullOrWhiteSpace(selectedItemText) Then
                    ' Add the full string toolstrip dropdown item to the bottom of the list with a seperator just above it
                    copyMenuItem.DropDownItems.Add(New ToolStripSeparator())
                    copyMenuItem.DropDownItems.Add(selectedItemText.Trim(), Nothing, AddressOf CopyToolStripItemTextToClipboard)
                End If

                If copyMenuItem.DropDownItems.Count > 0 Then
                    ' At least one dropdown item was built from the selected listview item so add the "Copy" menu item to the context menu strip
                    menuStrip.Items.Add(copyMenuItem)
                End If
            End If

            ' Add context menu items that should always be available
            menuStrip.Items.Add("Export All", Nothing, AddressOf BeginExportProcess)
            menuStrip.Items.Add("Refresh", Nothing, AddressOf LoaderBackgroundThread.RunWorkerAsync)

            AddHandler listView.ContextMenuStrip.Closed, AddressOf listView.ContextMenuStrip.Dispose
            listView.ContextMenuStrip = menuStrip
        End If
    End Sub

    Private Sub BeginExportProcess()
        ' Disable all tabs except the current one during the export process
        Dim allTabPagesExceptCurrent As List(Of TabPage) = OwnerTab.TabPages _
            .OfType(Of TabPage)() _
            .Where(Function(TabPage) Not TabPage.Text = Me.Text) _
            .ToList()

        For Each tabPage In allTabPagesExceptCurrent
            tabPage.Enabled = False
        Next

        ' Get information for setting the default export file name
        Dim selectedComputerName As String = Me.ComputerContext.OwnerForm.ResourceExplorer.SelectedNode.Text.Split(">")(1).Trim()
        Dim selectedTabName As String = Me.ComputerContext.LastSelectedTab.Text

        ' Prompt the user to select export file location
        Dim saveDialog As New SaveFileDialog() With
            {
                .InitialDirectory = "%UserProfile%",
                .FileName = String.Format("{0} - {1}", selectedComputerName, selectedTabName),
                .Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*",
                .RestoreDirectory = True
            }

        If saveDialog.ShowDialog() = DialogResult.OK Then
            Try
                ' The user chose the export file path. Now check if the file can safely be opened for writing.
                saveDialog.OpenFile().Close()

                ' Display the exporting progress bar control
                Dim exportProgress = NewProgressControl("Exporting...", Me.ExportingProgressBar, ProgressBarStyle.Continuous, Me.ExportBackgroundThread)
                Me.MainListView.InvokeAddControl(exportProgress)
                Me.MainListView.InvokeCenterControl()

                ' Delegate event handlers for the BackgroundWorker that will be performing the export.
                ' The DoWork() handler is to be delegated in the derived tab object's constructor.
                AddHandler Me.ExportBackgroundThread.ProgressChanged, AddressOf TabExportReportProgress
                AddHandler Me.ExportBackgroundThread.RunWorkerCompleted, Sub() TabExportCompleted(exportProgress, allTabPagesExceptCurrent)

                ' Start the export BackgroundWorker
                Me.ExportBackgroundThread.RunWorkerAsync(saveDialog.FileName)
            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            End Try
        Else
            ' User canceled so enable the previously disabled tabs
            For Each tabPage In allTabPagesExceptCurrent
                tabPage.Enabled = True
            Next
        End If
    End Sub

    Private Sub TabExportReportProgress(sender As Object, e As ProgressChangedEventArgs)
        ' ProgressChangedEventArgs.ProgressPercentage is currently unused, though it must be passed in the DoWork() implementation
        Me.ExportingProgressBar.Increment(1)
    End Sub

    Public Sub TabExportCompleted(exportingProgressBar As Control, allTabPagesExceptCurrent As List(Of TabPage))
        ' Remove the progress bar control from the user interface
        Me.MainListView.InvokeRemoveControl(exportingProgressBar)

        ' Enable the tabs that were disabled at the beginning of the export process
        For Each tabPage As TabPage In allTabPagesExceptCurrent
            tabPage.Enabled = True
        Next
    End Sub

    Public Sub ExportFromListView(listView As ListView, path As String)
        ' Create a data table for storing the export data
        Dim dataTable As New DataTable()
        ' Prepare the DataTable's schema
        Dim nonEmptyGroups = listView.Groups.Cast(Of ListViewGroup)().Where(Function(g) g.Items.Count > 0)
        nonEmptyGroups.ToList().ForEach(Function(group) dataTable.Columns.Add(group.Header))

        ' Get the total count of entries for calcuating progress
        Dim totalEntriesCount = nonEmptyGroups.SelectMany(Function(g) g.Items.Cast(Of ListViewItem)).Count()
        Dim currentEntryCount As Integer = 0

        ' Set the maximum value for the progress bar
        Me.UIThread(Sub() Me.ExportingProgressBar.Maximum = totalEntriesCount)

        ' Find the "maximum depth" of the available data. This number will end up being the number of rows in the DataTable.
        Dim maximumDepth As Integer = nonEmptyGroups.Max(Function(g) g.Items.Count)

        ' Begin looping to create each row in the data table
        For index As Integer = 0 To maximumDepth
            Dim row As DataRow = dataTable.NewRow()

            ' With each iteration loop through all listview groups
            For Each group As ListViewGroup In listView.Groups
                If Me.ExportBackgroundThread.CancellationPending Then Exit For

                If group.Items.Count > index Then
                    Dim currentItem = group.Items(index)

                    If currentItem.SubItems.Count > 0 Then
                        ' Concatenate all subitems to create a single entry
                        Dim entryText As String = String.Join(" ", currentItem.SubItems.Cast(Of ListViewSubItem)().Select(Function(subItem) subItem.Text.Trim()))

                        ' Apply the entry to the current row under the current group/column header
                        row(group.Header) = entryText

                        ' Report progress to the background worker
                        currentEntryCount += 1
                        Me.ExportBackgroundThread.ReportProgress(100 * (currentEntryCount / totalEntriesCount))
                    End If
                End If
            Next

            ' Add the fully populated row to the datatable
            dataTable.Rows.Add(row)
            If Me.ExportBackgroundThread.CancellationPending Then
                Exit For
            End If
        Next

        ' Write the list out to the CSV file chosen by the user in the base class Tab.BeginExportProcess()
        If dataTable.Rows.Count > 0 Then
            WriteDataTableToCSV(dataTable, path, includeHeader:=True)
        End If
    End Sub

    Public Sub AddToCurrentMenuStrip(toolStripItem As ToolStripItem)
        Dim menuItems = Me.MainListView.ContextMenuStrip.Items
        Dim selectedCount = Me.MainListView.SelectedItems.Count

        ' This will insert either above Copy, Refresh and seperator or above refresh
        Dim index As Integer = If(selectedCount > 0, menuItems.Count - 3, menuItems.Count - 1)
        menuItems.Insert(index, toolStripItem)
    End Sub

    Public Sub CopyToolStripItemTextToClipboard(sender As Object, e As EventArgs)
        Dim toolStripItem As ToolStripItem = sender
        If Not String.IsNullOrWhiteSpace(toolStripItem.Text) Then
            Clipboard.SetText(toolStripItem.Text)
        End If
    End Sub

    Public Function NewBaseListViewItems(masterListView As ListView, params As Object()) As ListViewItem()
        Dim items As New List(Of ListViewItem)

        If params IsNot Nothing Then
            For Each parameters() As Object In params
                Dim name As String = parameters(0)
                Dim groupName As String = parameters(1)
                Dim info As Object = parameters(2)

                Dim newItem As New ListViewItem() With
                    {
                        .Text = name,
                        .Group = masterListView.Groups(groupName)
                    }

                If Not info.GetType().IsArray Then
                    If TypeOf info Is String Then
                        If info Is Nothing Then Continue For
                        If info.ToString().ToUpper() = "LEGEND" Then Exit For
                    End If

                    If info Is Nothing AndAlso masterListView.Columns.Count <> 1 Then
                        Continue For
                    End If

                    If TypeOf info Is ManagementObject Then
                        newItem.Tag = info
                    Else
                        newItem.SubItems.Add(info)
                    End If
                Else
                    For Each subInfo As Object In info
                        If subInfo Is Nothing Then subInfo = " "
                        If TypeOf subInfo Is String AndAlso subInfo.ToString().ToUpper() = "LEGEND" Then
                            Exit For
                        End If

                        If TypeOf subInfo Is ManagementObject Then
                            newItem.Tag = subInfo
                            Continue For
                        End If

                        newItem.SubItems.Add(subInfo)
                    Next
                End If

                items.Add(newItem)
            Next
        End If

        Return items.ToArray()
    End Function

    Public Sub NewTabWriterItem(subject As String, body As Object, group As String)
        If body IsNot Nothing Then
            If TypeOf body Is String AndAlso String.IsNullOrWhiteSpace(body) Then
                Exit Sub
            End If

            If Me.TabWriterObjects Is Nothing Then
                Me.TabWriterObjects = New List(Of Object)()
            End If

            Me.TabWriterObjects.Add(New Object() {String.Format("   {0}", subject), group, body})
        End If
    End Sub

    Public Sub NewEnumWriterItem(subject As String, body As Object, group As String)
        If body IsNot Nothing Then
            If TypeOf body Is String AndAlso String.IsNullOrWhiteSpace(body) Then
                Exit Sub
            End If

            If Me.EnumWriterObjects Is Nothing Then
                Me.EnumWriterObjects = New List(Of Object)()
            End If

            Me.EnumWriterObjects.Add(New Object() {String.Format("   {0}", subject), group, body})
        End If
    End Sub

    Public Sub ShowListView(masterListView As ListView, splitContainerToDisplay As SplitContainer, panelToDisplay As Panels)
        Me.UIThread(Sub() masterListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent))

        Dim panel = splitContainerToDisplay.Controls(panelToDisplay - 1)
        panel.InvokeClearControls()
        panel.InvokeAddControl(masterListView)
    End Sub

    Public Sub ShowListView(masterListView As ListView, controlToDislay As Control)
        Me.UIThread(Sub() masterListView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.ColumnContent))

        ' Add the listview to the tab
        controlToDislay.InvokeClearControls()
        controlToDislay.InvokeAddControl(masterListView)
    End Sub

    Public Enum Panels
        Panel1 = 1
        Panel2 = 2
    End Enum

    Public Function ConvertDate(originalDate As String, Optional includeTime As Boolean = False) As String
        Dim convertedDate As String = Nothing

        Dim year As String
        Dim month As String
        Dim day As String

        If originalDate Is Nothing Then
            Return Nothing
        Else
            If originalDate.Contains("-") Then
                originalDate.Replace("-", String.Empty)
            End If

            year = CType(Mid(originalDate, 1, 4), Integer)
            month = CType(Mid(originalDate, 5, 2), Integer)
            day = CType(Mid(originalDate, 7, 2), Integer)
            convertedDate = String.Format("{0}/{1}/{2}", month, day, year)
        End If

        If includeTime Then
            convertedDate += String.Format(" {0}:{1}:{2}", Mid(originalDate, 9, 2), Mid(originalDate, 11, 2), Mid(originalDate, 13, 2))
        End If

        Return convertedDate
    End Function

    Public Sub Me_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If IsDisplayValidated() Then
            Me.LoaderBackgroundThread.RunWorkerAsync()
        End If
    End Sub

    Public Function IsDisplayValidated() As Boolean
        ' Does checks to see whether a background worker can begin enumeration on a computer.

        If Not Me.IsDisplayValidationNeeded Then
            If Not Me.LoaderBackgroundThread.IsBusy Then
                Return True
            Else
                Return False
            End If
        End If

        If Me.ComputerContext.OwnerForm.IsNodeSelected(Me.ComputerContext) Then
            If Me.Visible Then
                If Me.ComputerContext.ConnectionStatus = ComputerControl.ConnectionStatuses.Online AndAlso Me.ComputerContext.RespondsToPing() Then
                    If Not Me.LoaderBackgroundThread.IsBusy AndAlso Me.ComputerContext.LastSelectedTab IsNot Me Then
                        Me.ComputerContext.LastSelectedTab = Me
                        Return True
                    Else
                        Return False
                    End If
                Else
                    If Me.ComputerContext.ConnectionStatus <> ComputerControl.ConnectionStatuses.Online Then
                        If Not Me.LoaderBackgroundThread.IsBusy AndAlso Me.ComputerContext.LastSelectedTab IsNot Me Then
                            Me.ComputerContext.LastSelectedTab = Me
                            Return True
                        Else
                            Return False
                        End If
                    Else
                        Me.ComputerContext.SetConnectionStatus(ComputerControl.ConnectionStatuses.Offline)
                        Return False
                    End If

                End If
            Else
                Return False
            End If
        Else
            Return False
        End If
    End Function

    Public Function NewProgressControl(labelText As String, ByRef progressBar As ProgressBar, style As ProgressBarStyle, ByRef worker As BackgroundWorker) As Control
        ' Initialize "Progress" controls
        Dim label As New Label() With
            {
                .TextAlign = ContentAlignment.MiddleCenter,
                .Text = labelText,
                .AutoSize = False,
                .Height = 15,
                .Width = 100,
                .BackColor = Color.Transparent
            }

        ' Setting the properties for the progress bar
        progressBar = New ProgressBar() With
            {
                .Style = style,
                .Enabled = True,
                .Height = 10,
                .Width = 100
            }

        ' Setting the properties for the Cancel button
        Dim cancelButton As New Button() With
            {
                .Text = "Cancel",
                .Height = 30,
                .Width = 100
            }

        AddHandler cancelButton.Click, AddressOf worker.CancelAsync

        ' Setting the properties for the control that houses the progress bar and button
        Dim progress As New Control()
        With progress
            .Height = progressBar.Height + cancelButton.Height + label.Height + 10
            .Width = 110
            .BackColor = Color.White
            .InvokeAddControl(label)
            .Controls(0).Location = New Point(5, 5)
            .InvokeAddControl(progressBar)
            .Controls(1).Location = New Point(5, 20)
            .InvokeAddControl(cancelButton)
            .Controls(2).Location = New Point(5, 30)
        End With

        ' Draw a border around the control when it is being painted on the screen
        AddHandler progress.Paint, Sub(sender As Object, e As PaintEventArgs)
                                       ControlPaint.DrawBorder(e.Graphics, progress.ClientRectangle, Color.LightGray, ButtonBorderStyle.Solid)
                                   End Sub

        Return progress
    End Function

    Public Sub ShowTabLoaderProgress()
        Dim loaderProgress = NewProgressControl("Loading...", Me.LoadingProgressBar, ProgressBarStyle.Marquee, Me.LoaderBackgroundThread)
        Me.InvokeAddControl(loaderProgress)
        Me.InvokeCenterControl()
    End Sub

    Public Function UserCancellationPending() As Boolean
        If Me.LoaderBackgroundThread.CancellationPending Then
            Me.ComputerContext.WriteMessage("Stopping loading process!", Color.Red)
            Me.InvokeClearControls()

            Dim reloadButton As New Button() With {.Width = 100, .Text = "Reload..."}
            AddHandler reloadButton.Click, AddressOf Me.LoaderBackgroundThread.RunWorkerAsync
            Me.InvokeAddControl(reloadButton)
            Me.InvokeCenterControl()

            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetSelectedTabName() As String
        Return Me.OwnerTab.SelectedTab.Name
    End Function

    Public Function GetSelectedListViewItem(listView As ListView) As ListViewItem
        Dim item As ListViewItem = Nothing

        Me.UIThread(Sub()
                              If listView.SelectedItems.Count > 0 Then
                                  item = listView.SelectedItems(0)
                              End If
                          End Sub)

        Return item
    End Function

    Public Sub ValidateWMI()
        Try
            Dim wmiInstance = Me.ComputerContext.WMI.Query("SELECT LastBootUpTime FROM Win32_OperatingSystem")
            Dim lastBootUpTime As String = Me.ComputerContext.WMI.GetPropertyValue(wmiInstance, "LastBootUpTime")
            Dim convertedLastBootUpTime As String = ConvertDate(lastBootUpTime, includeTime:=True)

            If Me.ComputerContext.WMI.ConnectedTime < convertedLastBootUpTime Then
                Me.ComputerContext.WMI.Connect(Me.ComputerContext.Computer.Value, WMIController.ManagementScopes.All, async:=False)
            End If

        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))

            Try
                Me.ComputerContext.WMI.Connect(Me.ComputerContext.Computer.Value, WMIController.ManagementScopes.All, async:=False)
            Catch exc As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), exc.Message))
                Me.ComputerContext.SetConnectionStatus(ComputerControl.ConnectionStatuses.Offline)
            End Try
        End Try
    End Sub

    Public Function NewSplitContainer(splitterDistance As Integer) As SplitContainer
        Dim splitContainer As New SplitContainer() With
            {
                .Width = Me.Width,
                .Height = Me.Height - 25,
                .SplitterDistance = splitterDistance,
                .Anchor = AnchorStyles.Top + AnchorStyles.Bottom + AnchorStyles.Left + AnchorStyles.Right
            }

        Return splitContainer
    End Function

    Public Function NewSplitContainer(splitterDistance As Integer, panel2Label As String) As SplitContainer
        Dim serviceInfoLabel As New Label() With
            {
                .Dock = DockStyle.Fill,
                .TextAlign = ContentAlignment.MiddleCenter,
                .Text = panel2Label
            }

        Dim splitContainer As SplitContainer = NewSplitContainer(splitterDistance)
        splitContainer.Panel2.InvokeAddControl(serviceInfoLabel)

        Return splitContainer
    End Function

    Public Function NewSearchTextBox(text As String) As TextBox
        Dim searchTextBox As New TextBox() With
            {
                .Width = Me.Width,
                .Height = 25,
                .Anchor = AnchorStyles.Bottom + AnchorStyles.Left + AnchorStyles.Right,
                .Text = text,
                .Tag = text,
                .AutoCompleteSource = AutoCompleteSource.CustomSource,
                .AutoCompleteMode = AutoCompleteMode.SuggestAppend
            }

        ' Sift through the Main ListView
        Dim searchStrings As New AutoCompleteStringCollection()

        For Each item As ListViewItem In Me.MainListView.Items
            searchStrings.Add(GetListViewItemText(item))
        Next

        searchTextBox.AutoCompleteCustomSource = searchStrings

        AddHandler searchTextBox.GotFocus, AddressOf ClearSearchBox
        AddHandler searchTextBox.TextChanged, AddressOf OnSearchBoxTextChanged
        AddHandler searchTextBox.KeyDown, AddressOf OnSearchBoxKeyDown
        AddHandler Me.MainListView.KeyPress, AddressOf OnMainListViewKeyPress

        Me.SearchTextBox = searchTextBox

        Return Me.SearchTextBox
    End Function

    Public Sub ClearSearchBox()
        If GetSearchBoxText() = GetSearchBoxTag() Then
            Me.SearchTextBox.Clear()
        End If
    End Sub

    Public Sub OnSearchBoxTextChanged()
        Dim searchText = GetSearchBoxText()

        If Not String.IsNullOrWhiteSpace(searchText) Then
            If searchText <> GetSearchBoxTag() Then

                For Each item As ListViewItem In Me.MainListView.Items
                    If GetListViewItemText(item).Contains(searchText) Then
                        With Me.MainListView
                            .Items(item.Index).EnsureVisible()
                            .Items(item.Index).Selected = True
                            .Items(item.Index).Focused = True
                        End With

                        Exit Sub
                    End If
                Next
            End If
        End If
    End Sub

    Public Sub OnSearchBoxKeyDown(sender As Object, e As KeyEventArgs)
        If e.KeyCode = Keys.Up OrElse e.KeyCode = Keys.Down Then

            If Me.MainListView.SelectedItems.Count > 0 Then
                Dim selectedItem As ListViewItem = GetSelectedListViewItem(Me.MainListView)

                Select Case e.KeyCode
                    Case Keys.Up
                        If selectedItem.Index > 0 Then
                            Me.MainListView.Items(selectedItem.Index - 1).Selected = True
                            Me.MainListView.Items(selectedItem.Index - 1).Focused = True
                        End If

                    Case Keys.Down
                        If selectedItem.Index < Me.MainListView.Items.Count - 1 Then
                            Me.MainListView.Items(selectedItem.Index + 1).Selected = True
                            Me.MainListView.Items(selectedItem.Index + 1).Focused = True
                        End If
                End Select

                Me.MainListView.Focus()
            End If
        End If
    End Sub

    Public Sub OnMainListViewKeyPress(sender As Object, e As KeyPressEventArgs)
        If Char.IsLetterOrDigit(e.KeyChar) OrElse Char.IsSymbol(e.KeyChar) OrElse e.KeyChar = " " Then
            If GetSearchBoxText() = GetSearchBoxTag() Then
                Me.SearchTextBox.Text = String.Empty
            End If

            With Me.SearchTextBox
                .Text += e.KeyChar
                .SelectionStart = .Text.Length - 1
                .ScrollToCaret()
                .Focus()
            End With
        End If

        If e.KeyChar = ControlChars.Back Then
            If Not String.IsNullOrWhiteSpace(Me.SearchTextBox.Text) Then
                With Me.SearchTextBox
                    .Text = .Text.Substring(0, .Text.Length - 2)
                    .SelectionStart = .Text.Length - 1
                    .ScrollToCaret()
                    .Focus()
                End With
            End If
        End If
    End Sub

    Private Function GetSearchBoxText() As String
        Return Me.SearchTextBox.Text.Trim().ToUpper()
    End Function

    Private Function GetSearchBoxTag() As String
        Dim text = String.Empty
        If TypeOf Me.SearchTextBox.Tag Is String Then
            text = Me.SearchTextBox.Tag.Trim().ToUpper()
        End If

        Return text
    End Function

    Private Function GetListViewItemText(item As ListViewItem) As String
        Return item.Text.Trim().ToUpper()
    End Function

    Public Sub ClearEnumeratorVars()
        Me.TabWriterObjects = Nothing
        Me.MainListView = Nothing
        Me.IsMainListViewCreated = False
        Me.SearchTextBox = Nothing
    End Sub

    Public Sub Me_PreviewKeyDown(sender As Object, e As PreviewKeyDownEventArgs) Handles Me.PreviewKeyDown
        If e.KeyCode = Keys.F5 Then
            Me.LoaderBackgroundThread.RunWorkerAsync()
        End If
    End Sub

End Class