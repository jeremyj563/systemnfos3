Imports System.ComponentModel
Imports System.IO
Imports System.Threading
Imports System.Reflection
Imports System.Runtime.InteropServices
Imports System.Net
Imports System.Net.NetworkInformation

Public Class TaskResultControl

    Public Property OwnerForm As MainForm

    Private Property ComputerNames As List(Of String)
    Private Property Task As String
    Private Property Args As String
    Private Property IsViewable As Boolean
    Private Property SortingColumn As ColumnHeader

    Public Sub New(computerNames As List(Of String), task As String, args As String, isViewable As Boolean, ownerForm As MainForm)
        InitializeComponent()
        Me.ComputerNames = computerNames
        Me.Task = task
        Me.Args = args
        Me.IsViewable = isViewable
        Me.OwnerForm = ownerForm
    End Sub

    Private Sub TaskResultControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim resultsMenuStrip As New ContextMenuStrip
        resultsMenuStrip.Items.Add("Get Info", Nothing, AddressOf GetInfo)
        ResultsListView.ContextMenuStrip = resultsMenuStrip

        Dim taskWorker As New BackgroundWorker() With {.WorkerSupportsCancellation = True}
        AddHandler taskWorker.DoWork, AddressOf PrepareTask
        taskWorker.RunWorkerAsync()
    End Sub

    Private Sub PrepareTask(sender As Object, e As DoWorkEventArgs)
        For Each computerName As String In Me.ComputerNames
            Dim psexecWorker As New BackgroundWorker()
            AddHandler psexecWorker.DoWork, AddressOf Me.PsExecWork
            psexecWorker.RunWorkerAsync(computerName)
        Next
    End Sub

    Private Sub PsExecWork(sender As Object, e As DoWorkEventArgs)
        Dim computerName As String = e.Argument.ToString()
        Dim viewable As String = Nothing
        Dim files As String() = Me.Task.Split("\")
        Dim fileName As String = files(files.Length)
        Dim appPath As String = Path.GetDirectoryName(Me.Task)

        If Me.IsViewable Then
            viewable = " -i "
        Else
            viewable = " "
        End If

        If Me.Task.Contains(" ") Then
            Me.Task = String.Format("""{0}""", Me.Task)
        End If

        If appPath.Contains(" ") Then
            appPath = String.Format("""{0}""", appPath)
        End If

        If fileName.Contains(" ") Then
            fileName = String.Format("""{0}""", fileName)
        End If

        PerformListViewWork(computerName, "Initializing...")
        PerformListViewWork(computerName, "Connecting to client...")

        Dim psExec As New Process()
        If RespondsToPing(computerName) Then
            Try
                Dim psExecProcess As New ProcessStartInfo With
                {
                    .FileName = My.Settings.PsExecPath,
                    .Arguments = String.Format("-accepteula -s \\{0} robocopy.exe {1} {2}", computerName, appPath & fileName, ITFolderLocalPath),
                    .WindowStyle = ProcessWindowStyle.Hidden
                }

                If Me.Task.ToUpper().Trim().EndsWith(".PRINTEREXPORT""") OrElse Me.Task.ToUpper().Trim().EndsWith(".PRINTEREXPORT") Then
                    PerformListViewWork(computerName, "Copying printer export file...")

                    If psExec.ExitCode <> 0 AndAlso psExec.ExitCode <> 1 Then
                        PerformListViewWork(computerName, String.Format("Robocopy process failed! Exit Code: {0}", psExec.ExitCode))
                        Exit Sub
                    End If

                    Thread.Sleep(1000)
                    Dim filePath As String = String.Format("""{0}""", Path.Combine(ITFolderLocalPath, files(files.Length)))

                    psExecProcess.Arguments = String.Format("-s \\{0} {1} -r -f {2} -O FORCE", computerName, Path.Combine(Environment.SystemDirectory, "spool\tools\PrintBrm.exe"), filePath)
                    If Not String.IsNullOrWhiteSpace(Args) Then
                        psExecProcess.Arguments += String.Format(" {0}", Args)
                    End If
                Else
                    psExecProcess.Arguments = String.Format("-accepteula -s {0}\\{1} {2}", viewable, computerName, Task)
                    If Not String.IsNullOrWhiteSpace(Args) Then
                        psExecProcess.Arguments += String.Format(" {0}", Args)
                    End If
                End If

                PerformListViewWork(computerName, String.Format("Running {0}\{1}", files(files.Length - 1), files(files.Length)))
                psExec = Process.Start(psExecProcess)
                psExec.WaitForExit()
                PerformListViewWork(computerName, String.Format("Process Complete! Exit Code: {0}", psExec.ExitCode.ToString()))

            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                PerformListViewWork(computerName, String.Format("Process Incomplete! Error Message: {0}", ex.Message))
            Finally
                psExec.Close()
            End Try
        Else
            PerformListViewWork(computerName, "Process Incomplete! Unable To Ping")
        End If
    End Sub

    Private Sub PerformListViewWork(computerName As String, status As String)
        Me.UIThread(Sub()
                        If status = "Initializing..." Then
                            Dim item As New ListViewItem() With {.Text = computerName}
                            item.SubItems.Add(status)
                            Me.ResultsListView.Items.Add(item)
                        Else
                            For Each item As ListViewItem In Me.ResultsListView.Items
                                If item.Text.ToUpper() = computerName.ToUpper() Then
                                    item.SubItems(1).Text = status
                                End If
                            Next
                        End If
                    End Sub)
    End Sub

    Public Sub ReleaseObject([object] As Object)
        Try
            Marshal.ReleaseComObject([object])
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        Finally
            [object] = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub lvSysStatus_ColumnClick(sender As Object, e As ColumnClickEventArgs)
        Dim newSortingColumn As ColumnHeader = ResultsListView.Columns(e.Column)
        Dim sortOrder As SortOrder

        If Me.SortingColumn Is Nothing Then
            sortOrder = SortOrder.Ascending
        Else
            If newSortingColumn.Equals(Me.SortingColumn) Then
                If Me.SortingColumn.Text.StartsWith("> ") Then
                    sortOrder = SortOrder.Descending
                Else
                    sortOrder = SortOrder.Ascending
                End If
            Else
                sortOrder = SortOrder.Ascending
            End If
            Me.SortingColumn.Text = Me.SortingColumn.Text.Substring(2)
        End If

        Me.SortingColumn = newSortingColumn
        If sortOrder = SortOrder.Ascending Then
            Me.SortingColumn.Text = String.Format("> {0}", Me.SortingColumn.Text)
        Else
            Me.SortingColumn.Text = String.Format("< {0}", Me.SortingColumn.Text)
        End If

        ResultsListView.ListViewItemSorter = New ListViewComparer(e.Column, sortOrder)
        ResultsListView.Sort()
    End Sub

    Private Sub GetInfo()
        For Each item As ListViewItem In Me.ResultsListView.SelectedItems
            Dim searcher As New DataSourceSearcher(item.Text, Me.OwnerForm.BindingSource)
            Dim results As List(Of Computer) = searcher.GetComputers()

            If results IsNot Nothing Then
                Me.OwnerForm.UserInputComboBox.SelectedItem = results.First()
                Me.OwnerForm.SubmitButton.PerformClick()
            End If
        Next
    End Sub

    Private Function RespondsToPing(computerName As String) As Boolean
        Try
            Dim ipAddress As IPAddress = Dns.GetHostAddresses(computerName).FirstOrDefault()
            If ipAddress IsNot Nothing Then
                Dim ping As New Ping()
                Dim pingResult As Boolean = ping.Send(ipAddress, 500).Status = IPStatus.Success

                Return pingResult
            End If
        Catch ex As Exception
            ' Silently fail
        End Try

        Return False
    End Function

    Private Class ListViewComparer
        Implements IComparer

        Private Property ColumnNumber As Integer
        Private Property SortOrder As SortOrder

        Public Sub New(columnNumber As Integer, sortOrder As SortOrder)
            Me.ColumnNumber = columnNumber
            Me.SortOrder = sortOrder
        End Sub

        Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
            Dim itemX As ListViewItem = x
            Dim ItemY As ListViewItem = y

            Dim stringX As String
            If itemX.SubItems.Count <= Me.ColumnNumber Then
                stringX = String.Empty
            Else
                stringX = itemX.SubItems(Me.ColumnNumber).Text
            End If

            Dim stringY As String
            If ItemY.SubItems.Count <= Me.ColumnNumber Then
                stringY = String.Empty
            Else
                stringY = ItemY.SubItems(Me.ColumnNumber).Text
            End If

            If Me.SortOrder = SortOrder.Ascending Then
                If IsNumeric(stringX) AndAlso IsNumeric(stringY) Then
                    Return Val(stringX).CompareTo(Val(stringY))
                ElseIf IsDate(stringX) AndAlso IsDate(stringY) Then
                    Return Date.Parse(stringX).CompareTo(Date.Parse(stringY))
                Else
                    Return String.Compare(stringX, stringY)
                End If
            Else
                If IsNumeric(stringX) AndAlso IsNumeric(stringY) Then
                    Return Val(stringY).CompareTo(Val(stringX))
                ElseIf IsDate(stringX) AndAlso IsDate(stringY) Then
                    Return Date.Parse(stringY).CompareTo(Date.Parse(stringX))
                Else
                    Return String.Compare(stringY, stringX)
                End If
            End If
        End Function

    End Class

End Class
