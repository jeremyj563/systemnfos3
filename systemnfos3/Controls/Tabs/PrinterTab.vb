Imports System.ComponentModel
Imports System.Reflection
Imports Microsoft.Win32

Public Class PrinterTab
    Inherits Tab

    Private Property PrinterInfoListView As ListView
    Private Property DefaultPrinter As String() = Nothing

    Public Sub New(ownerTab As TabControl, computerContext As ComputerControl)
        MyBase.New(ownerTab, computerContext)

        Me.Text = "Printers"
        AddHandler MyBase.LoaderBackgroundThread.DoWork, AddressOf Me.InitializePrinterTab
        AddHandler MyBase.ExportBackgroundThread.DoWork, AddressOf Me.ExportPrinterInfo
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

        Me.PrinterInfoListView = NewBasicInfoListView(4)
        With Me.PrinterInfoListView
            .MultiSelect = True
            .Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPR), ListViewGroups.lsvgPR))
            .Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPD), ListViewGroups.lsvgPD))
            .Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgPO), ListViewGroups.lsvgPO))
        End With

        ' Find the default printer
        Me.DefaultPrinter = Nothing
        Dim currentUserRegistryKey As String = Nothing
        Dim registry As New RegistryController(ComputerContext.WMI.X86Scope)

        Select Case ComputerContext.UserLoggedOn

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
                    LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                End Try
            Case Else
        End Select

        If Not String.IsNullOrWhiteSpace(currentUserRegistryKey) Then
            Try
                Me.DefaultPrinter = registry.GetKeyValue(String.Format("{0}\Software\Microsoft\Windows NT\CurrentVersion\Windows", currentUserRegistryKey), "Device", RegistryController.RegistryKeyValueTypes.String, RegistryHive.Users).ToString().Split(",")
            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
            End Try
        End If

        ' Create the Legend for Printers
        Dim printers As ManagementObjectCollection = ComputerContext.WMI.Query("SELECT Name, PortName, PrintProcessor, PrintJobDataType, CreationClassName FROM Win32_Printer")
        If printers IsNot Nothing AndAlso printers.Count > 0 Then
            Try
                NewTabWriterItem("Printer Name:", New String() {"Printer Port Name:", "Print Processor:", "Data Type:", "LEGEND"}, NameOf(ListViewGroups.lsvgPR))
                For Each printer As ManagementBaseObject In printers
                    If MyBase.UserCancellationPending() Then Exit Sub

                    If Me.DefaultPrinter IsNot Nothing AndAlso printer.Properties("name").Value = Me.DefaultPrinter(0) Then
                        NewTabWriterItem(String.Format("{0} (Default Printer)", printer.Properties("Name").Value), New Object() {printer.Properties("PortName").Value, printer.Properties("PrintProcessor").Value, printer.Properties("PrintJobDataType").Value, printer}, NameOf(ListViewGroups.lsvgPR))
                    Else
                        NewTabWriterItem(printer.Properties("Name").Value, New Object() {printer.Properties("PortName").Value, printer.Properties("PrintProcessor").Value, printer.Properties("PrintJobDataType").Value, printer}, NameOf(ListViewGroups.lsvgPR))
                    End If
                Next
            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                NewTabWriterItem("Error:", New Object() {ex.Message, "LEGEND"}, NameOf(ListViewGroups.lsvgPR))
            End Try
        End If

        ' Create the legend for Printer Drivers
        Dim printerDrivers As ManagementObjectCollection = ComputerContext.WMI.Query("SELECT Name, DriverPath, CreationClassName FROM Win32_PrinterDriver")
        If printerDrivers IsNot Nothing AndAlso printerDrivers.Count > 0 Then
            Try
                NewTabWriterItem("Driver Name:", New String() {"Driver Path:", "LEGEND"}, NameOf(ListViewGroups.lsvgPD))
                For Each driver As ManagementBaseObject In printerDrivers
                    If MyBase.UserCancellationPending() Then Exit Sub

                    NewTabWriterItem(driver.Properties("Name").Value, New Object() {driver.Properties("DriverPath").Value, driver}, NameOf(ListViewGroups.lsvgPD))
                Next
            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                NewTabWriterItem("Error:", New Object() {ex.Message, "LEGEND"}, NameOf(ListViewGroups.lsvgPD))
            End Try
        End If

        ' Create the legend for TCP/IP Printer Ports
        Dim tcpIPPorts As ManagementObjectCollection = ComputerContext.WMI.Query("SELECT Name, HostAddress, PortNumber, CreationClassName FROM Win32_TCPIPPrinterPort")
        If tcpIPPorts IsNot Nothing AndAlso tcpIPPorts.Count > 0 Then
            Try
                NewTabWriterItem("Port Name:", New String() {"Host Address:", "Port Number:", "LEGEND"}, NameOf(ListViewGroups.lsvgPO))
                For Each port As ManagementBaseObject In tcpIPPorts
                    If MyBase.UserCancellationPending() Then Exit Sub

                    NewTabWriterItem(port.Properties("Name").Value, New Object() {port.Properties("HostAddress").Value, port.Properties("PortNumber").Value, port}, NameOf(ListViewGroups.lsvgPO))
                Next
            Catch ex As Exception
                LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                NewTabWriterItem("Error:", New Object() {ex.Message, "LEGEND"}, NameOf(ListViewGroups.lsvgPO))
            End Try
        End If

        If Me.TabWriterObjects IsNot Nothing Then
            Me.PrinterInfoListView.Items.AddRange(NewBaseListViewItems(Me.PrinterInfoListView, Me.TabWriterObjects.ToArray()))
        End If

        AddHandler Me.MainListView.ContextMenuStripChanged, AddressOf AddToPrinterMenuStrip
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

        Select Case Me.MainListView.SelectedItems.Count

            Case 1
                ' One Printer is selected. Show All options. Ignore the Legends
                If Me.MainListView.SelectedItems(0).Tag IsNot Nothing Then
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                    Select Case Me.MainListView.SelectedItems(0).Tag.Properties("CreationClassName").Value
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
                If Me.MainListView.SelectedItems.Count > 0 Then
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                    AddToCurrentMenuStrip(New ToolStripMenuItem("Delete Printer Objects", Nothing, AddressOf DeleteAnyPrinterType))
                    AddToCurrentMenuStrip(New ToolStripSeparator)
                End If
        End Select
    End Sub

    Private Sub AddPrinter()
        Dim printerAdd As New RemoteTools(RemoteTools.RemoteTools.PrinterAdd, Me.ComputerContext)
        AddHandler printerAdd.WorkCompleted, AddressOf LoaderBackgroundThread.RunWorkerAsync
        printerAdd.BeginWork()
    End Sub

    Private Sub AddPrinterCab()
        Dim printerAddCab As New RemoteTools(RemoteTools.RemoteTools.PrinterAddCab, Me.ComputerContext)
        AddHandler printerAddCab.WorkCompleted, AddressOf LoaderBackgroundThread.RunWorkerAsync
        printerAddCab.BeginWork()
    End Sub

    Private Sub NewPrinterCab()
        Dim printerNewCab As New RemoteTools(RemoteTools.RemoteTools.PrinterNewCab, Me.ComputerContext)
        printerNewCab.BeginWork()
    End Sub

    Private Sub RenamePrinter(sender As Object, e As EventArgs)
        Dim printerRename As New RemoteTools(RemoteTools.RemoteTools.PrinterRename, Me.ComputerContext, New ManagementObject() {GetSelectedListViewItem(MainListView).Tag})
        AddHandler printerRename.WorkCompleted, AddressOf Me.LoaderBackgroundThread.RunWorkerAsync
        printerRename.BeginWork()
    End Sub

    Private Sub SetDefaultPrinter()
        Dim printerSetDefault As New RemoteTools(RemoteTools.RemoteTools.PrinterSetDefault, Me.ComputerContext, New ManagementObject() {GetSelectedListViewItem(MainListView).Tag})
        AddHandler printerSetDefault.WorkCompleted, AddressOf Me.LoaderBackgroundThread.RunWorkerAsync
        printerSetDefault.BeginWork()
    End Sub

    Private Sub OpenPrinterQueue()
        Dim printerOpenQueue As New RemoteTools(RemoteTools.RemoteTools.PrinterOpenQueue, Me.ComputerContext, New ManagementObject() {GetSelectedListViewItem(MainListView).Tag})
        printerOpenQueue.BeginWork()
    End Sub

    Private Sub OpenPrinterProperties()
        Dim printerOpenProperties As New RemoteTools(RemoteTools.RemoteTools.PrinterOpenProperties, Me.ComputerContext, New ManagementObject() {GetSelectedListViewItem(MainListView).Tag})
        printerOpenProperties.BeginWork()
    End Sub

    Private Sub EditPrinterPort()
        Dim printerEditPort As New RemoteTools(RemoteTools.RemoteTools.PrinterEditPort, Me.ComputerContext, New ManagementObject() {GetSelectedListViewItem(MainListView).Tag})
        AddHandler printerEditPort.WorkCompleted, AddressOf LoaderBackgroundThread.RunWorkerAsync
        printerEditPort.BeginWork()
    End Sub

    Private Sub DeleteAnyPrinterType(sender As Object, e As EventArgs)
        If Me.MainListView.SelectedItems.Count > 0 Then
            Dim printersToDelete As New List(Of ManagementObject)()
            For Each printer As ListViewItem In Me.MainListView.SelectedItems
                If printer.Tag Is Nothing Then Continue For

                printersToDelete.Add(printer.Tag)
            Next

            If printersToDelete IsNot Nothing Then
                Dim printerDeleteAnyType As New RemoteTools(RemoteTools.RemoteTools.PrinterDeleteAnyType, Me.ComputerContext, printersToDelete.ToArray())
                AddHandler printerDeleteAnyType.WorkCompleted, AddressOf LoaderBackgroundThread.RunWorkerAsync

                printerDeleteAnyType.BeginWork()
            End If
        End If
    End Sub

    Private Sub ExportPrinterInfo(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.PrinterInfoListView, userSelectedCSVFilePath)
    End Sub

End Class
