Imports System.ComponentModel
Imports System.Reflection
Imports Microsoft.Win32

Public Class PrinterTab
    Inherits BaseTab

    Private Property PrinterInfoListView As ListView
    Private Property DefaultPrinter As String() = Nothing

    Public Sub New(ownerTab As TabControl, computerPanel As ComputerPanel)
        MyBase.New(ownerTab, computerPanel)

        Me.Text = "Printers"
        AddHandler MyBase.InitWorker.DoWork, AddressOf Me.InitializePrinterTab
        AddHandler MyBase.ExportWorker.DoWork, AddressOf Me.ExportPrinterInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgPR As String = "Printers"
        Public Shared ReadOnly Property lsvgPD As String = "Printer Drivers"
        Public Shared ReadOnly Property lsvgPO As String = "TCP/IP Ports"
    End Structure

    Private Sub InitializePrinterTab()
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Me.PrinterInfoListView = NewBaseListView(4)
        With Me.PrinterInfoListView
            .MultiSelect = True
            .Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPR), ListViewGroups.lsvgPR))
            .Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPD), ListViewGroups.lsvgPD))
            .Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPO), ListViewGroups.lsvgPO))
        End With

        ' Find the default printer
        Me.DefaultPrinter = Nothing
        Dim currentUserRegistryKey As String = Nothing
        Dim registry As New RegistryController(ComputerPanel.WMI.X86Scope)

        Select Case ComputerPanel.UserLoggedOn

            Case "True"
                Try
                    For Each user In registry.GetKeyValues(String.Empty, RegistryHive.Users)
                        If MyBase.UserCancellationPending() Then Exit Sub

                        Select Case user
                            Case ".DEFAULT"
                                Continue For
                            Case "S-1-5-18"
                                Continue For
                            Case "S-1-5-19"
                                Continue For
                            Case "S-1-5-20"
                                Continue For
                            Case Else
                                For Each keyValue In registry.GetKeyValues(user, RegistryHive.Users)
                                    If MyBase.UserCancellationPending() Then Exit Sub

                                    If keyValue = "Volatile Environment" Then
                                        currentUserRegistryKey = user
                                        Exit Select
                                    End If
                                Next
                        End Select

                        If Not String.IsNullOrWhiteSpace(currentUserRegistryKey) Then
                            Exit For
                        End If
                    Next
                Catch ex As Exception
                    LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
                End Try
            Case Else
        End Select

        If Not String.IsNullOrWhiteSpace(currentUserRegistryKey) Then
            Try
                Dim keyPath = $"{currentUserRegistryKey}\Software\Microsoft\Windows NT\CurrentVersion\Windows"
                Me.DefaultPrinter = registry.GetKeyValue(keyPath, "Device", RegistryController.RegistryKeyValueTypes.String, RegistryHive.Users).ToString().Split(",")
            Catch ex As Exception
                LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
            End Try
        End If

        ' Create the Legend for Printers
        Dim printers As ManagementObjectCollection = ComputerPanel.WMI.Query("SELECT Name, PortName, PrintProcessor, PrintJobDataType, CreationClassName FROM Win32_Printer")
        If printers IsNot Nothing AndAlso printers.Count > 0 Then
            Try
                AddTabWriterItem("Printer Name:", New String() {"Printer Port Name:", "Print Processor:", "Data Type:", "LEGEND"}, NameOf(ListViewGroups.lsvgPR))
                For Each printer As ManagementBaseObject In printers
                    If MyBase.UserCancellationPending() Then Exit Sub

                    If Me.DefaultPrinter IsNot Nothing AndAlso printer.Properties("name").Value = Me.DefaultPrinter(0) Then
                        AddTabWriterItem($"{printer.Properties("Name").Value} (Default Printer)", New Object() {printer.Properties("PortName").Value, printer.Properties("PrintProcessor").Value, printer.Properties("PrintJobDataType").Value, printer}, NameOf(ListViewGroups.lsvgPR))
                    Else
                        AddTabWriterItem(printer.Properties("Name").Value, New Object() {printer.Properties("PortName").Value, printer.Properties("PrintProcessor").Value, printer.Properties("PrintJobDataType").Value, printer}, NameOf(ListViewGroups.lsvgPR))
                    End If
                Next
            Catch ex As Exception
                LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
                AddTabWriterItem("Error:", New Object() {ex.Message, "LEGEND"}, NameOf(ListViewGroups.lsvgPR))
            End Try
        End If

        ' Create the legend for Printer Drivers
        Dim printerDrivers As ManagementObjectCollection = ComputerPanel.WMI.Query("SELECT Name, DriverPath, CreationClassName FROM Win32_PrinterDriver")
        If printerDrivers IsNot Nothing AndAlso printerDrivers.Count > 0 Then
            Try
                AddTabWriterItem("Driver Name:", New String() {"Driver Path:", "LEGEND"}, NameOf(ListViewGroups.lsvgPD))
                For Each driver As ManagementBaseObject In printerDrivers
                    If MyBase.UserCancellationPending() Then Exit Sub

                    AddTabWriterItem(driver.Properties("Name").Value, New Object() {driver.Properties("DriverPath").Value, driver}, NameOf(ListViewGroups.lsvgPD))
                Next
            Catch ex As Exception
                LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
                AddTabWriterItem("Error:", New Object() {ex.Message, "LEGEND"}, NameOf(ListViewGroups.lsvgPD))
            End Try
        End If

        ' Create the legend for TCP/IP Printer Ports
        Dim tcpIPPorts As ManagementObjectCollection = ComputerPanel.WMI.Query("SELECT Name, HostAddress, PortNumber, CreationClassName FROM Win32_TCPIPPrinterPort")
        If tcpIPPorts IsNot Nothing AndAlso tcpIPPorts.Count > 0 Then
            Try
                AddTabWriterItem("Port Name:", New String() {"Host Address:", "Port Number:", "LEGEND"}, NameOf(ListViewGroups.lsvgPO))
                For Each port As ManagementBaseObject In tcpIPPorts
                    If MyBase.UserCancellationPending() Then Exit Sub

                    AddTabWriterItem(port.Properties("Name").Value, New Object() {port.Properties("HostAddress").Value, port.Properties("PortNumber").Value, port}, NameOf(ListViewGroups.lsvgPO))
                Next
            Catch ex As Exception
                LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
                AddTabWriterItem("Error:", New Object() {ex.Message, "LEGEND"}, NameOf(ListViewGroups.lsvgPO))
            End Try
        End If

        If Me.TabWriterObjects IsNot Nothing Then
            Me.PrinterInfoListView.Items.AddRange(NewBaseListViewItems(Me.PrinterInfoListView, Me.TabWriterObjects.ToArray()))
        End If

        AddHandler Me.CurrentListView.ContextMenuStripChanged, AddressOf AddToPrinterMenuStrip
        ShowListView(Me.PrinterInfoListView, Me)
    End Sub

    Private Sub AddToPrinterMenuStrip()
        Dim installMenuItem As New ToolStripMenuItem("Install")
        With installMenuItem.DropDownItems
            .Add("Single Printer", Nothing, AddressOf AddPrinter)
            .Add("Printer Export File", Nothing, AddressOf AddPrinterCab)
        End With

        AddToCurrentMenuStrip(installMenuItem)
        AddToCurrentMenuStrip(New ToolStripSeparator)
        AddToCurrentMenuStrip(New ToolStripMenuItem("Create Printer Export File", Nothing, AddressOf NewPrinterCab))

        Select Case Me.CurrentListView.SelectedItems.Count

            Case 1
                ' One Printer is selected. Show All options. Ignore the Legends
                If Me.CurrentListView.SelectedItems(0).Tag IsNot Nothing Then
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                    Select Case Me.CurrentListView.SelectedItems(0).Tag.Properties("CreationClassName").Value
                        Case "Win32_Printer"
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Rename Printer", Nothing, AddressOf RenamePrinter))
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Set as Default", Nothing, AddressOf SetDefaultPrinter))
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Printer Queue", Nothing, AddressOf OpenPrinterQueue))
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Printer Properties", Nothing, AddressOf OpenPrinterProperties))
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Change Port", Nothing, AddressOf EditPrinterPort))
                            AddToCurrentMenuStrip(New ToolStripSeparator)
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Delete Printer", Nothing, AddressOf DeleteAnyPrinterType))
                            AddToCurrentMenuStrip(New ToolStripSeparator)
                        Case "Win32_PrinterDriver"
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Delete Driver", Nothing, AddressOf DeleteAnyPrinterType))
                            AddToCurrentMenuStrip(New ToolStripSeparator)
                        Case "Win32_TCPIPPrinterPort"
                            AddToCurrentMenuStrip(New ToolStripMenuItem("Delete Port", Nothing, AddressOf DeleteAnyPrinterType))
                            AddToCurrentMenuStrip(New ToolStripSeparator)
                    End Select
                End If

            Case Else
                ' Multiple printers are selected. Only show the delete options
                If Me.CurrentListView.SelectedItems.Count > 0 Then
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Delete Printer Objects", Nothing, AddressOf DeleteAnyPrinterType))
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                End If
        End Select
    End Sub

    Private Sub AddPrinter()
        Dim printerAdd As New RemoteTools(RemoteTools.RemoteTools.PrinterAdd, Me.ComputerPanel)
        AddHandler printerAdd.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        printerAdd.BeginWork()
    End Sub

    Private Sub AddPrinterCab()
        Dim printerAddCab As New RemoteTools(RemoteTools.RemoteTools.PrinterAddCab, Me.ComputerPanel)
        AddHandler printerAddCab.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        printerAddCab.BeginWork()
    End Sub

    Private Sub NewPrinterCab()
        Dim printerNewCab As New RemoteTools(RemoteTools.RemoteTools.PrinterNewCab, Me.ComputerPanel)
        printerNewCab.BeginWork()
    End Sub

    Private Sub RenamePrinter(sender As Object, e As EventArgs)
        Dim printerRename As New RemoteTools(RemoteTools.RemoteTools.PrinterRename, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(CurrentListView).Tag})
        AddHandler printerRename.WorkCompleted, AddressOf Me.InitWorker.RunWorkerAsync
        printerRename.BeginWork()
    End Sub

    Private Sub SetDefaultPrinter()
        Dim printerSetDefault As New RemoteTools(RemoteTools.RemoteTools.PrinterSetDefault, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(CurrentListView).Tag})
        AddHandler printerSetDefault.WorkCompleted, AddressOf Me.InitWorker.RunWorkerAsync
        printerSetDefault.BeginWork()
    End Sub

    Private Sub OpenPrinterQueue()
        Dim printerOpenQueue As New RemoteTools(RemoteTools.RemoteTools.PrinterOpenQueue, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(CurrentListView).Tag})
        printerOpenQueue.BeginWork()
    End Sub

    Private Sub OpenPrinterProperties()
        Dim printerOpenProperties As New RemoteTools(RemoteTools.RemoteTools.PrinterOpenProperties, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(CurrentListView).Tag})
        printerOpenProperties.BeginWork()
    End Sub

    Private Sub EditPrinterPort()
        Dim printerEditPort As New RemoteTools(RemoteTools.RemoteTools.PrinterEditPort, Me.ComputerPanel, New ManagementObject() {GetSelectedListViewItem(CurrentListView).Tag})
        AddHandler printerEditPort.WorkCompleted, AddressOf InitWorker.RunWorkerAsync
        printerEditPort.BeginWork()
    End Sub

    Private Sub DeleteAnyPrinterType(sender As Object, e As EventArgs)
        If Me.CurrentListView.SelectedItems.Count > 0 Then
            Dim printersToDelete As New List(Of ManagementObject)()
            For Each printer As ListViewItem In Me.CurrentListView.SelectedItems
                If printer.Tag Is Nothing Then Continue For

                printersToDelete.Add(printer.Tag)
            Next

            If printersToDelete IsNot Nothing Then
                Dim printerDeleteAnyType As New RemoteTools(RemoteTools.RemoteTools.PrinterDeleteAnyType, Me.ComputerPanel, printersToDelete.ToArray())
                AddHandler printerDeleteAnyType.WorkCompleted, AddressOf InitWorker.RunWorkerAsync

                printerDeleteAnyType.BeginWork()
            End If
        End If
    End Sub

    Private Sub ExportPrinterInfo(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.PrinterInfoListView, userSelectedCSVFilePath)
    End Sub

End Class
