Imports System.ComponentModel
Imports System.Net
Imports System.Net.NetworkInformation
Imports System.Reflection
Imports System.Runtime.InteropServices

Public Class QueryResultControl
    Public Property OwnerForm As MainForm

    Private Property ComputerNames As List(Of String)
    Private Property Properties As List(Of String)
    Private Property SortingColumn As ColumnHeader

    Public Sub New(computerNames As List(Of String), properties As List(Of String), ownerForm As MainForm)
        InitializeComponent()
        Me.ComputerNames = computerNames
        Me.Properties = properties
        Me.OwnerForm = ownerForm
    End Sub

    Private Sub QueryResultControl_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        For Each [property] As String In Me.Properties
            AddColumn(Split([property], ">>").First)
        Next

        Dim resultsMenuStrip As New ContextMenuStrip()
        resultsMenuStrip.Items.Add("Get Info", Nothing, AddressOf GetInfo)

        Me.ResultsListView.ContextMenuStrip = resultsMenuStrip

        Dim queryWorker As New BackgroundWorker() With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}
        queryWorker.RunWorkerAsync()
        AddHandler queryWorker.DoWork, AddressOf AddComputerNameToListView
    End Sub

    Private Sub AddComputerNameToListView(sender As Object, e As DoWorkEventArgs)
        Dim psexecWorker As New BackgroundWorker() With {.WorkerReportsProgress = True, .WorkerSupportsCancellation = True}

        For Each computerName As String In Me.ComputerNames
            AddHandler psexecWorker.DoWork, AddressOf PsExecWork
            psexecWorker.RunWorkerAsync(computerName)
        Next
    End Sub

    Private Sub PsExecWork(sender As Object, e As DoWorkEventArgs)
        Dim computerName As String = e.Argument.ToString()

        Dim wmiQueryWork As New WMIQueryWork()
        PerformListViewWork(computerName, "Initializing...", 1)

        For Each parentProperty In Me.Properties
            Dim [property] As String = parentProperty.Split(">>").First()
            Dim serviceName As String = [property].Split(">>").First()
            Dim index As Integer = GetSubItemIndex(serviceName)
            Select Case [property]
                Case "LDAP"
                    PerformListViewWork(computerName, "Querying Data...", index)
                    PerformListViewWork(computerName, DetermineActiveDirectoryAttribute(serviceName, computerName), index)

                Case "WMI"
                    PerformListViewWork(computerName, "Pinging...", index)
                    If RespondsToPing(computerName) Then
                        PerformListViewWork(computerName, "Querying Data...", index)
                        PerformListViewWork(computerName, wmiQueryWork.GetWMIResult(serviceName, computerName), index)
                    Else
                        PerformListViewWork(computerName, "Unable To Ping", index)
                    End If
            End Select
        Next
    End Sub

    Private Sub PerformListViewWork(computerName As String, status As String, index As Integer)
        Me.UIThread(Sub()
                        Me.ResultsListView.BeginUpdate()

                        If status.ToUpper().Contains("INITIALIZING") Then
                            Dim listViewItem As New ListViewItem() With {.Text = computerName}
                            For i As Integer = 0 To Me.ResultsListView.Columns.Count - 1
                                listViewItem.SubItems.Add(status)
                            Next

                            Me.ResultsListView.Items.Add(listViewItem)
                        Else
                            For Each computerItem As ListViewItem In Me.ResultsListView.Items
                                If computerItem.Text.ToUpper() = computerName.ToUpper() Then
                                    computerItem.SubItems(index).Text = status
                                End If
                            Next
                        End If

                        Me.ResultsListView.EndUpdate()
                    End Sub)
    End Sub

    Private Sub AddColumn(Name As String)
        Dim column As New ColumnHeader() With
            {
                .Name = Name,
                .Text = Name,
                .Width = 120
            }

        Me.ResultsListView.Columns.Add(column)
    End Sub

    Private Sub ReleaseObject([object] As Object)
        Try
            Marshal.ReleaseComObject([object])
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        Finally
            [object] = Nothing
            GC.Collect()
        End Try
    End Sub

    Private Sub ResultsListView_ColumnClick(sender As Object, e As ColumnClickEventArgs) Handles ResultsListView.ColumnClick
        Dim newSortingColumn As ColumnHeader = ResultsListView.Columns(e.Column)
        Dim order As SortOrder

        If Me.SortingColumn Is Nothing Then
            order = SortOrder.Ascending
        Else
            If newSortingColumn.Equals(Me.SortingColumn) Then
                If Me.SortingColumn.Text.StartsWith("> ") Then
                    order = SortOrder.Descending
                Else
                    order = SortOrder.Ascending
                End If
            Else
                order = SortOrder.Ascending
            End If

            Me.SortingColumn.Text = Me.SortingColumn.Text.Substring(2)
        End If

        Me.SortingColumn = newSortingColumn
        If order = SortOrder.Ascending Then
            Me.SortingColumn.Text = String.Format("> {0}", Me.SortingColumn.Text)
        Else
            Me.SortingColumn.Text = String.Format("< {0}", Me.SortingColumn.Text)
        End If

        Me.ResultsListView.ListViewItemSorter = New ListViewComparer(e.Column, order)
        Me.ResultsListView.Sort()
    End Sub

    Private Function DetermineActiveDirectoryAttribute(criteria As String, computerName As String) As String
        Select Case criteria
            Case "NAC IP Address"
                Return Split(GetActiveDirectoryAttribute("networkAddress", computerName), ",")(0)
            Case "NAC MAC Address"
                Return Split(GetActiveDirectoryAttribute("networkAddress", computerName), ",")(1)
            Case "OS Name"
                Return GetActiveDirectoryAttribute("operatingSystem", computerName)
            Case "OS Description"
                Return GetActiveDirectoryAttribute("description", computerName)
            Case "OS Version"
                Return GetActiveDirectoryAttribute("operatingSystemVersion", computerName)
            Case "CS Model"
                Return GetActiveDirectoryAttribute("info", computerName)
            Case "CS Serial Number"
                Return GetActiveDirectoryAttribute("carLicense", computerName)
            Case "CS Bios Version"
                Return GetActiveDirectoryAttribute("roomNumber", computerName)
            Case "CS Image Date"
                Return GetActiveDirectoryAttribute("co", computerName)
            Case "CS Location"
                Return GetActiveDirectoryAttribute("physicalDeliveryOfficeName", computerName)
            Case "CS Last Logged On User"
                Return GetActiveDirectoryAttribute("uid", computerName)
        End Select

        Return Nothing
    End Function

    Private Function GetActiveDirectoryAttribute([property] As String, computerName As String) As String
        Try
            Dim entry As New DirectoryEntry()
            entry.AuthenticationType = AuthenticationTypes.Secure

            Dim searcher As New DirectorySearcher(entry)
            With searcher
                .Filter = String.Format("(&(objectClass=computer)(|(name={0})))", computerName)
                .PropertiesToLoad.Add([property])
            End With

            Dim result As SearchResult = searcher.FindOne()
            If result.Properties([property]).Count > 0 Then
                Return result.Properties([property])(0).ToString()
            End If
        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try

        Return Nothing
    End Function

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

    Private Function GetSubItemIndex(subItemName As String) As Integer
        For Each column As ColumnHeader In Me.ResultsListView.Columns
            If subItemName = column.Text Then Return column.Index
        Next

        Return Nothing
    End Function

    Private Sub GetInfo()
        For Each listViewItem As ListViewItem In Me.ResultsListView.SelectedItems
            Dim searcher As New DataSourceSearcher(listViewItem.Text, Me.OwnerForm.BoundData)

            Dim results As List(Of Computer) = searcher.GetComputers()
            If results IsNot Nothing Then
                Me.OwnerForm.UserInputComboBox.SelectedItem = results.First()
                Me.OwnerForm.SubmitButton.PerformClick()
            End If
        Next
    End Sub

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
            Dim itemY As ListViewItem = y

            Dim stringX As String
            If itemX.SubItems.Count <= ColumnNumber Then
                stringX = String.Empty
            Else
                stringX = itemX.SubItems(ColumnNumber).Text
            End If

            Dim stringY As String
            If itemY.SubItems.Count <= ColumnNumber Then
                stringY = String.Empty
            Else
                stringY = itemY.SubItems(ColumnNumber).Text
            End If

            If Me.SortOrder = SortOrder.Ascending Then
                If IsNumeric(stringX) And IsNumeric(stringY) Then
                    Return Val(stringX).CompareTo(Val(stringY))
                ElseIf IsDate(stringX) And IsDate(stringY) Then
                    Return Date.Parse(stringX).CompareTo(Date.Parse(stringY))
                Else
                    Return String.Compare(stringX, stringY)
                End If
            Else
                If IsNumeric(stringX) And IsNumeric(stringY) Then
                    Return Val(stringY).CompareTo(Val(stringX))
                ElseIf IsDate(stringX) And IsDate(stringY) Then
                    Return Date.Parse(stringY).CompareTo(Date.Parse(stringX))
                Else
                    Return String.Compare(stringY, stringX)
                End If
            End If

        End Function

    End Class

End Class
