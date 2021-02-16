Imports System.ComponentModel

Public Class AdvancedTab
    Inherits Tab

    Private Property advancedInfoListView As ListView

    Public Sub New(ownerTab As TabControl, computerPanel As ComputerPanel)
        MyBase.New(ownerTab, computerPanel)

        Me.Text = "Advanced Information"
        AddHandler MyBase.InitWorker.DoWork, AddressOf Me.InitializeAdvancedTab
        AddHandler MyBase.ExportWorker.DoWork, AddressOf Me.ExportAdvancedInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgOS As String = "Operating System"
        Public Shared ReadOnly Property lsvgCM As String = "Computer Model"
        Public Shared ReadOnly Property lsvgBI As String = "BIOS"
        Public Shared ReadOnly Property lsvgCP As String = "CPU"
        Public Shared ReadOnly Property lsvgNI As String = "Network Adapter"
        Public Shared ReadOnly Property lsvgHD As String = "Hard Drive"
        Public Shared ReadOnly Property lsvgMI As String = "Memory"
    End Structure

    Private Sub InitializeAdvancedTab()
        ' Clear the tab of all child controls
        Me.InvokeClearControls()
        ShowTabLoaderProgress()
        ValidateWMI()
        ClearEnumeratorVars()

        Me.advancedInfoListView = NewBasicInfoListView(2)
        With Me.advancedInfoListView.Groups
            ' Create the ListView Groups
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgOS), ListViewGroups.lsvgOS))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgCM), ListViewGroups.lsvgCM))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgBI), ListViewGroups.lsvgBI))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgCP), ListViewGroups.lsvgCP))
        End With

        LoadBaseAdvancedInfo()
        LoadNetInterfaceInstances()
        LoadHDDInterfaceInstances()
        LoadMemoryInterfaceInstances()

        Dim copyMenuStrip As New ContextMenuStrip()
        copyMenuStrip.Items.Add("Copy")
        Me.advancedInfoListView.ContextMenuStrip = copyMenuStrip

        ' Add Items into ListView
        Me.advancedInfoListView.Items.AddRange(MyBase.NewBaseListViewItems(Me.advancedInfoListView, MyBase.TabWriterObjects.ToArray()))
        ShowListView(Me.advancedInfoListView, Me)
    End Sub

    Private Sub LoadBaseAdvancedInfo()
        NewTabWriterItem("Name:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Caption FROM Win32_OperatingSystem"), "Caption"), NameOf(ListViewGroups.lsvgOS))

        NewTabWriterItem("Service Pack:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT CSDVersion FROM Win32_OperatingSystem"), "CSDVersion"), NameOf(ListViewGroups.lsvgOS))
        NewTabWriterItem("Architecture:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT AddressWidth FROM Win32_Processor"), "AddressWidth"), NameOf(ListViewGroups.lsvgOS))
        NewTabWriterItem("OS Install Date:", ConvertDate(Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT InstallDate FROM Win32_OperatingSystem"), "InstallDate"), True), NameOf(ListViewGroups.lsvgOS))
        NewTabWriterItem("Last Boot:", ConvertDate(Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT LastBootUpTime FROM Win32_OperatingSystem"), "LastBootUpTime"), True), NameOf(ListViewGroups.lsvgOS))

        NewTabWriterItem("Manufacturer:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Manufacturer FROM Win32_ComputerSystem"), "Manufacturer"), NameOf(ListViewGroups.lsvgCM))
        NewTabWriterItem("Model:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Model FROM Win32_ComputerSystem"), "Model"), NameOf(ListViewGroups.lsvgCM))
        NewTabWriterItem("Serial Number:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT IdentifyingNumber FROM Win32_ComputerSystemProduct"), "IdentifyingNumber"), NameOf(ListViewGroups.lsvgCM))
        NewTabWriterItem("Type:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Name FROM Win32_BIOS"), "Name"), NameOf(ListViewGroups.lsvgBI))

        NewTabWriterItem("Version:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT SMBIOSBIOSVersion FROM Win32_BIOS"), "SMBIOSBIOSVersion"), NameOf(ListViewGroups.lsvgBI))
        NewTabWriterItem("Type:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Name FROM Win32_Processor"), "Name"), NameOf(ListViewGroups.lsvgCP))
        NewTabWriterItem("Cores:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT NumberOfCores FROM Win32_Processor"), "NumberOfCores"), NameOf(ListViewGroups.lsvgCP))
    End Sub

    Private Sub LoadNetInterfaceInstances()
        ' Network Interfaces
        Me.advancedInfoListView.Groups.Add(NameOf(ListViewGroups.lsvgNI), ListViewGroups.lsvgNI)
        Dim networkInterfaceCount As Integer = 0
        For Each instance In Me.ComputerPanel.WMI.Query($"SELECT Description, IPAddress, MACAddress, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE DNSDomain='{My.Settings.DomainName}'")
            If MyBase.UserCancellationPending() Then Exit Sub

            networkInterfaceCount += 1

            Dim adapterName As String = instance.Properties("Description").Value
            Me.advancedInfoListView.Groups.Add(NameOf(ListViewGroups.lsvgNI) & networkInterfaceCount, $"{ListViewGroups.lsvgNI} ({adapterName})")

            For index = 0 To instance.GetPropertyValue("IPAddress").Length - 1
                NewTabWriterItem("IP Address:", instance.GetPropertyValue("IPAddress")(index), NameOf(ListViewGroups.lsvgNI) & networkInterfaceCount)
            Next

            NewTabWriterItem("MAC Address:", instance.Properties("MACAddress").Value, NameOf(ListViewGroups.lsvgNI) & networkInterfaceCount)

            If instance.GetPropertyValue("DefaultIPGateway") IsNot Nothing Then
                For index = 0 To instance.GetPropertyValue("DefaultIPGateway").Length - 1
                    NewTabWriterItem("Default Gateway:", instance.GetPropertyValue("DefaultIPGateway")(index), NameOf(ListViewGroups.lsvgNI) & networkInterfaceCount)
                Next
            End If
        Next
    End Sub

    Private Sub LoadHDDInterfaceInstances()
        ' Hard Disk Interfaces
        Me.advancedInfoListView.Groups.Add(NameOf(ListViewGroups.lsvgHD), ListViewGroups.lsvgHD)
        Dim hddInterfaceCount As Integer = 0
        For Each hddInterface As ManagementBaseObject In Me.ComputerPanel.WMI.Query("SELECT Name, FreeSpace, Size, VolumeSerialNumber, VolumeName FROM Win32_LogicalDisk WHERE DriveType = 3")
            If MyBase.UserCancellationPending() Then Exit Sub

            hddInterfaceCount += 1

            Dim driveLetter As String = hddInterface.Properties("Name").Value
            Me.advancedInfoListView.Groups.Add(New ListViewGroup(NameOf(ListViewGroups.lsvgHD) & hddInterfaceCount, $"{ListViewGroups.lsvgHD} ({driveLetter})"))

            NewTabWriterItem("Volume Name:", hddInterface.Properties("VolumeName").Value, NameOf(ListViewGroups.lsvgHD) & hddInterfaceCount)
            NewTabWriterItem("Serial Number:", hddInterface.Properties("VolumeSerialNumber").Value, NameOf(ListViewGroups.lsvgHD) & hddInterfaceCount)
            NewTabWriterItem("Total Space (GBs):", Math.Round(hddInterface.Properties("Size").Value / (1024 ^ 3), 2), NameOf(ListViewGroups.lsvgHD) & hddInterfaceCount)
            NewTabWriterItem("Free Space (GBs):", Math.Round(hddInterface.Properties("FreeSpace").Value / (1024 ^ 3), 2), NameOf(ListViewGroups.lsvgHD) & hddInterfaceCount)
        Next
    End Sub

    Private Sub LoadMemoryInterfaceInstances()
        ' Memory Interfaces
        Me.advancedInfoListView.Groups.Add(NameOf(ListViewGroups.lsvgMI), ListViewGroups.lsvgMI)

        Dim memoryInterfaceCount As Integer = 0
        Dim totalCapacity As Long = 0

        For Each memoryInterface As ManagementBaseObject In Me.ComputerPanel.WMI.Query("SELECT Capacity, Speed, SerialNumber, PartNumber FROM Win32_PhysicalMemory")
            If MyBase.UserCancellationPending() Then Exit Sub

            memoryInterfaceCount += 1

            Dim capacity As Long = memoryInterface.Properties("Capacity").Value / 1073741824
            Me.advancedInfoListView.Groups.Add(NameOf(ListViewGroups.lsvgMI) & memoryInterfaceCount, $"{ListViewGroups.lsvgMI} {memoryInterfaceCount}")

            NewTabWriterItem("Capacity:", $"{capacity} GB(s)", NameOf(ListViewGroups.lsvgMI) & memoryInterfaceCount)
            NewTabWriterItem("Speed:", $"{memoryInterface.Properties("Speed").Value} MHz", NameOf(ListViewGroups.lsvgMI) & memoryInterfaceCount)
            NewTabWriterItem("Serial Number:", memoryInterface.Properties("SerialNumber").Value, NameOf(ListViewGroups.lsvgMI) & memoryInterfaceCount)
            NewTabWriterItem("Part Number:", memoryInterface.Properties("PartNumber").Value, NameOf(ListViewGroups.lsvgMI) & memoryInterfaceCount)
            totalCapacity += capacity
        Next

        NewTabWriterItem("Total Memory Capacity:", $"{totalCapacity} GB(s)", NameOf(ListViewGroups.lsvgMI))
    End Sub

    Private Sub ExportAdvancedInfo(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.advancedInfoListView, userSelectedCSVFilePath)
    End Sub

End Class