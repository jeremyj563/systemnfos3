Imports System.ComponentModel
Imports Microsoft.Win32

Public Class MainTab
    Inherits Tab

    Private Property BasicInfoListView As ListView
    Private Property ConnectionStatus As ComputerPanel.ConnectionStatuses
    Private Property RunOnlyOnce As Boolean = False

    Public Sub New(ownerTab As TabControl, computerPanel As ComputerPanel, connectionStatus As ComputerPanel.ConnectionStatuses)
        MyBase.New(ownerTab, computerPanel)

        Me.Text = "Basic Information"
        Me.ConnectionStatus = connectionStatus
        AddHandler MyBase.InitWorker.DoWork, AddressOf Me.InitializeMainTab
        AddHandler MyBase.ExportWorker.DoWork, AddressOf Me.ExportBasicInfo
    End Sub

    Private Structure ListViewGroups
        Public Shared ReadOnly Property lsvgOS As String = "Operating System"
        Public Shared ReadOnly Property lsvgCI As String = "Computer"
        Public Shared ReadOnly Property lsvgEN As String = "Encryption"
        Public Shared ReadOnly Property lsvgNI As String = "Network"
        Public Shared ReadOnly Property lsvgUI As String = "User"
    End Structure

    Private Sub InitializeMainTab()
        Me.InvokeClearControls()
        ShowTabLoaderProgress()

        If Me.ConnectionStatus = ComputerPanel.ConnectionStatuses.Online Then
            MyBase.ValidateWMI()
        End If

        ClearEnumeratorVars()

        Me.BasicInfoListView = NewBasicInfoListView(2)
        With Me.BasicInfoListView.Groups
            ' Create the ListView Groups
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgOS), ListViewGroups.lsvgOS))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgCI), ListViewGroups.lsvgCI))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgEN), ListViewGroups.lsvgEN))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgNI), ListViewGroups.lsvgNI))
            .Add(New ListViewGroup(NameOf(ListViewGroups.lsvgUI), ListViewGroups.lsvgUI))
        End With

        Select Case Me.ConnectionStatus
            Case ComputerPanel.ConnectionStatuses.Online
                LoadOnlineStatusComputers()
            Case Else
                LoadOtherStatusComputers()
        End Select

        Me.BasicInfoListView.Items.AddRange(NewBaseListViewItems(Me.BasicInfoListView, Me.TabWriterObjects.ToArray()))

        ShowListView(Me.BasicInfoListView, Me)
        Me.RunOnlyOnce = True
    End Sub

    Private Sub LoadOnlineStatusComputers()
        NewTabWriterItem("Operating System Name:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Caption FROM Win32_OperatingSystem"), "Caption"), NameOf(ListViewGroups.lsvgOS))
        NewTabWriterItem("Architecture:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT AddressWidth FROM Win32_Processor"), "AddressWidth"), NameOf(ListViewGroups.lsvgOS))
        NewTabWriterItem("Model:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Model FROM Win32_ComputerSystem"), "Model"), NameOf(ListViewGroups.lsvgCI))
        NewTabWriterItem("Serial Number:", Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT IdentifyingNumber FROM Win32_ComputerSystemProduct"), "IdentifyingNumber"), NameOf(ListViewGroups.lsvgCI))

        If Me.ComputerPanel.Computer.ActiveDirectoryContainer.LastLogonExtensionAttribute IsNot Nothing Then
            NewTabWriterItem("Last Logon:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.LastLogonExtensionAttribute, NameOf(ListViewGroups.lsvgCI))
        End If

        If Me.ComputerPanel.Computer.ActiveDirectoryContainer.PhysicalDeliveryOfficeName IsNot Nothing Then
            NewTabWriterItem("Location:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.PhysicalDeliveryOfficeName, NameOf(ListViewGroups.lsvgCI))
        End If

        NewTabWriterItem("Distinguished Name:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("distinguishedName"), NameOf(ListViewGroups.lsvgCI))

        If Convert.ToBoolean(Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("userAccountControl") And 2) Then
            NewTabWriterItem("Container Status:", "Disabled", NameOf(ListViewGroups.lsvgCI))
            If Not Me.RunOnlyOnce Then
                Dim message = "NOTE: Computer is Disabled. Contact Imaging for assistance."
                Me.ComputerPanel.WriteMessage(message, Color.OrangeRed)
            End If
        Else
            NewTabWriterItem("Container Status:", "Enabled", NameOf(ListViewGroups.lsvgCI))
        End If

        LoadTPMInstances()
        LoadBitLockerInstances()
        LoadNetAdapterInstances()
        LoadLoggedOnUser()
    End Sub

    Private Sub LoadOtherStatusComputers()
        ' Code for loading offline/degraded/slow status computers
        NewTabWriterItem("Operating System Name:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("OperatingSystem"), NameOf(ListViewGroups.lsvgOS))
        NewTabWriterItem("Model:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("info"), NameOf(ListViewGroups.lsvgCI))
        NewTabWriterItem("Serial Number:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("carLicense"), NameOf(ListViewGroups.lsvgCI))

        Dim lastLogon As String = Me.ComputerPanel.Computer.ActiveDirectoryContainer.LastLogonExtensionAttribute
        If Not String.IsNullOrWhiteSpace(lastLogon) Then
            NewTabWriterItem("Last Logon:", lastLogon, NameOf(ListViewGroups.lsvgCI))
        End If

        If Me.ComputerPanel.Computer.ActiveDirectoryContainer.PhysicalDeliveryOfficeName IsNot Nothing Then
            NewTabWriterItem("Location:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.PhysicalDeliveryOfficeName, NameOf(ListViewGroups.lsvgCI))
        End If

        If Convert.ToBoolean(Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("userAccountControl") And 2) Then
            NewTabWriterItem("Container Status:", "Disabled", NameOf(ListViewGroups.lsvgCI))

            If Not Me.RunOnlyOnce Then
                Dim message = "NOTE: Computer is Disabled. Contact Imaging for assistance."
                Me.ComputerPanel.WriteMessage(message, Color.OrangeRed)
            End If
        Else
            NewTabWriterItem("Container Status:", "Enabled", NameOf(ListViewGroups.lsvgCI))
        End If

        NewTabWriterItem("Distinguished Name:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("distinguishedName"), NameOf(ListViewGroups.lsvgCI))

        Dim networkAddresses As String = Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("networkAddress")
        If Not String.IsNullOrWhiteSpace(networkAddresses) AndAlso networkAddresses.Split(",").Length > 0 Then
            NewTabWriterItem("IP Address:", networkAddresses.Split(",")(0), NameOf(ListViewGroups.lsvgNI))
            NewTabWriterItem("MAC Address:", networkAddresses.Split(",")(1), NameOf(ListViewGroups.lsvgNI))
        End If

        Dim uid As Object = Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("uid")
        If uid IsNot Nothing Then
            Dim user As New LDAPContainer("user", "sAMAccountName", uid)
            NewTabWriterItem("Username:", user.GetAttribute("sAMAccountName"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Display Name:", user.GetAttribute("displayName"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Phone Number:", user.GetAttribute("telephoneNumber"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Location:", user.GetAttribute("streetAddress"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Office Symbol:", user.GetAttribute("physicalDeliveryOfficeName"), NameOf(ListViewGroups.lsvgUI))
        End If

        If Me.TabWriterObjects Is Nothing Then
            NewTabWriterItem("Error:", "Unable to enumerate offline data", Nothing)
        End If
    End Sub

    Private Sub LoadTPMInstances()
        If Me.ComputerPanel.WMI.TPMScope.IsConnected Then
            Try
                Dim tpm As New ManagementClass(Me.ComputerPanel.WMI.TPMScope, Me.ComputerPanel.WMI.TPMScope.Path, New ObjectGetOptions)
                Dim tpmInstances = tpm.GetInstances()

                If tpmInstances IsNot Nothing AndAlso tpmInstances.Count > 0 Then
                    For Each tpmInstance As ManagementObject In tpmInstances
                        If MyBase.UserCancellationPending() Then Exit Sub

                        Dim enableStatus(0) As Object
                        tpmInstance.InvokeMethod("IsEnabled", enableStatus)

                        Dim activateStatus(0) As Object
                        tpmInstance.InvokeMethod("IsActivated", activateStatus)

                        Select Case enableStatus(0)
                            Case True
                                NewTabWriterItem("TPM Status:", "On", NameOf(ListViewGroups.lsvgEN))
                            Case False
                                NewTabWriterItem("TPM Status:", "Off", NameOf(ListViewGroups.lsvgEN))
                        End Select

                        Select Case activateStatus(0)
                            Case True
                                NewTabWriterItem("TPM Activation:", "Activated", NameOf(ListViewGroups.lsvgEN))
                            Case False
                                NewTabWriterItem("TPM Activation:", "Not Activated", NameOf(ListViewGroups.lsvgEN))
                        End Select
                    Next
                Else
                    NewTabWriterItem("TPM Status", "Off", NameOf(ListViewGroups.lsvgEN))
                End If
            Catch ex As Exception
                ' Fail silently
            End Try
        End If
    End Sub

    Private Sub LoadBitLockerInstances()
        If Me.ComputerPanel.WMI.BitLockerScope.IsConnected Then
            Try
                Dim bitLocker As ManagementClass = New ManagementClass(Me.ComputerPanel.WMI.BitLockerScope, Me.ComputerPanel.WMI.BitLockerScope.Path, New ObjectGetOptions)
                Dim bitLockerInstances = bitLocker.GetInstances()

                If bitLockerInstances IsNot Nothing AndAlso bitLockerInstances.Count > 0 Then
                    Dim queryText = "SELECT DeviceID FROM Win32_Volume WHERE BootVolume = 1 AND DriveType = 3"
                    Dim bootDeviceID As String = Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query(queryText), "DeviceID")

                    For Each bitLockerInstance As ManagementObject In bitLockerInstances
                        If MyBase.UserCancellationPending() Then Exit Sub

                        Dim driveDeviceID As String = bitLockerInstance.Properties("DeviceID").Value
                        If driveDeviceID = bootDeviceID Then

                            Dim protectionStatus(0) As Object
                            bitLockerInstance.InvokeMethod("GetProtectionStatus", protectionStatus)
                            Dim hddEncryptionStatus(1) As Object
                            bitLockerInstance.InvokeMethod("GetConversionStatus", hddEncryptionStatus)

                            NewTabWriterItem("Bitlocker Status:", "Enabled", NameOf(ListViewGroups.lsvgEN))

                            Select Case protectionStatus(0)
                                Case 0
                                    NewTabWriterItem("Protection Status:", "Off", NameOf(ListViewGroups.lsvgEN))
                                Case 1
                                    NewTabWriterItem("Protection Status:", "On", NameOf(ListViewGroups.lsvgEN))
                            End Select

                            Select Case hddEncryptionStatus(0)
                                Case 0
                                    NewTabWriterItem("Encryption Status:", "Fully Decrypted", NameOf(ListViewGroups.lsvgEN))
                                Case 1
                                    NewTabWriterItem("Encryption Status:", "Fully Encrypted", NameOf(ListViewGroups.lsvgEN))
                                Case 2
                                    NewTabWriterItem("Encryption Status:", "Encrypting", NameOf(ListViewGroups.lsvgCI))
                                    NewTabWriterItem("Amount:", hddEncryptionStatus(1) & "%", NameOf(ListViewGroups.lsvgEN))
                                Case 3
                                    NewTabWriterItem("Encryption Status:", "Decrypting", NameOf(ListViewGroups.lsvgCI))
                                    NewTabWriterItem("Amount:", hddEncryptionStatus(1) & "%", NameOf(ListViewGroups.lsvgEN))
                                Case 4
                                    NewTabWriterItem("Encryption Status:", "Encrypting Paused", NameOf(ListViewGroups.lsvgCI))
                                    NewTabWriterItem("Amount:", hddEncryptionStatus(1) & "%", NameOf(ListViewGroups.lsvgEN))
                                Case 5
                                    NewTabWriterItem("Encryption Status:", "Decrypting Paused", NameOf(ListViewGroups.lsvgCI))
                                    NewTabWriterItem("Amount:", hddEncryptionStatus(1) & "%", NameOf(ListViewGroups.lsvgEN))
                            End Select

                            Exit For
                        End If
                    Next
                Else
                    NewTabWriterItem("BitLocker Status:", "Disabled", NameOf(ListViewGroups.lsvgEN))
                End If
            Catch ex As Exception
                ' Fail silently
            End Try
        End If
    End Sub

    Private Sub LoadNetAdapterInstances()
        Dim queryText = $"SELECT Description, IPAddress, MACAddress, DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE DNSDomain='{My.Settings.DomainName}'"
        Dim networkAdapters As ManagementObjectCollection = Me.ComputerPanel.WMI.Query(queryText)

        If networkAdapters IsNot Nothing Then
            Dim count As Integer = 0
            For Each networkAdapter As ManagementBaseObject In networkAdapters
                If MyBase.UserCancellationPending() Then Exit Sub

                count += 1
                If count >= 2 Then
                    NewTabWriterItem(" ", " ", NameOf(ListViewGroups.lsvgNI))
                End If

                NewTabWriterItem("Adapter Description:", networkAdapter.Properties("Description").Value, NameOf(ListViewGroups.lsvgNI))
                NewTabWriterItem("IP Address:", networkAdapter.GetPropertyValue("IPAddress")(0), NameOf(ListViewGroups.lsvgNI))
                NewTabWriterItem("MAC Address:", networkAdapter.Properties("MACAddress").Value, NameOf(ListViewGroups.lsvgNI))
            Next
        End If
    End Sub

    Private Sub LoadLoggedOnUser()
        ' Find out who is logged into a computer
        Dim registry As RegistryController = New RegistryController(Me.ComputerPanel.WMI.X86Scope)
        Dim userName As String = Nothing

        ' Check to see if explorer.exe is open.
        Dim explorerQueryResult As String = Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Name FROM Win32_Process WHERE Name = 'explorer.exe'"), "Name")

        ' If explorer is open then...
        If Not String.IsNullOrWhiteSpace(explorerQueryResult) AndAlso explorerQueryResult.ToUpper() = "EXPLORER.EXE" Then
            ' A user is logged in
            Me.ComputerPanel.UserLoggedOn = True
            ' What does WMI say the logged in user is?
            userName = Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT UserName FROM Win32_ComputerSystem"), "UserName")

            ' If WMI says there is an error getting the username then...
            If String.IsNullOrWhiteSpace(userName) Then
                ' Look in the registry keys under HKU hive
                For Each keyValue As String In registry.GetKeyValues(String.Empty, RegistryHive.Users)
                    If MyBase.UserCancellationPending() Then Exit Sub

                    ' And then check every key for the volatile environment folder
                    For Each nextSubKey In registry.GetKeyValues(keyValue, RegistryHive.Users)
                        If MyBase.UserCancellationPending() Then Exit Sub

                        ' If it finds the volatile environment folder then...
                        If nextSubKey.ToUpper() = "VOLATILE ENVIRONMENT" Then
                            ' Find out what the registry says for the logged in user
                            userName = registry.GetKeyValue($"{keyValue}\{nextSubKey}", "UserName", RegistryController.RegistryKeyValueTypes.String, RegistryHive.Users)

                            ' But if the registry says nothing then
                            If String.IsNullOrWhiteSpace(userName) Then
                                ' I give up... best guess is the last person that logged in is the current user. We can grab that name from AD since we record it there.
                                If Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("uid") IsNot Nothing Then
                                    NewTabWriterItem("Current User:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("uid"), NameOf(ListViewGroups.lsvgUI))
                                End If
                            Else
                                ' If the registry gives me an answer, then I know who is logged in
                                NewTabWriterItem("Current User:", userName, NameOf(ListViewGroups.lsvgUI))
                            End If
                        End If
                    Next
                Next
            Else
                ' Or if WMI was reliable the entire time, then I trust WMI :)
                NewTabWriterItem("Current User:", userName, NameOf(ListViewGroups.lsvgUI))
            End If
        Else
            Me.ComputerPanel.UserLoggedOn = False
        End If

        ' Since a user is logged in
        If Me.ComputerPanel.UserLoggedOn Then
            ' I will let the technician know
            NewTabWriterItem("Logged On:", "Yes", NameOf(ListViewGroups.lsvgUI))

            Dim logonUI As String = Me.ComputerPanel.WMI.GetPropertyValue(Me.ComputerPanel.WMI.Query("SELECT Name FROM Win32_Process WHERE Name = 'logonui.exe'"), "Name")
            If Not String.IsNullOrWhiteSpace(logonUI) AndAlso logonUI.ToUpper() = "LOGONUI.EXE" Then
                NewTabWriterItem("Away From Keyboard:", "Yes", NameOf(ListViewGroups.lsvgUI))
            Else
                NewTabWriterItem("Away From Keyboard:", "No", NameOf(ListViewGroups.lsvgUI))
            End If
        Else
            ' OR since a user is not logged in...
            ' Lets get the last logged in username from ldap
            userName = Me.ComputerPanel.Computer.ActiveDirectoryContainer.GetAttribute("uid")

            If Not String.IsNullOrWhiteSpace(userName) Then
                NewTabWriterItem("Last Logged In User:", userName, NameOf(ListViewGroups.lsvgUI))
            End If
            ' And then tell the technician no one is logged into the computer
            NewTabWriterItem("Logged On:", "No", NameOf(ListViewGroups.lsvgUI))
        End If

        If Not String.IsNullOrWhiteSpace(userName) Then
            ' Username was found so get data from the ldap 'user' class instance

            If userName.Contains("\") Then
                ' Username came back in sAMAccountName format ie "domainName\userName".
                ' We need to strip off the domain name since we only want the username for performing the LDAP query below.
                userName = userName.Split("\")(1)
            End If

            Dim user As New LDAPContainer("user", "sAMAccountName", userName)
            NewTabWriterItem("Display Name:", user.GetAttribute("displayName"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Last Logon:", Me.ComputerPanel.Computer.ActiveDirectoryContainer.LastLogonExtensionAttribute, NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Phone Number:", user.GetAttribute("telephoneNumber"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Location:", user.GetAttribute("streetAddress"), NameOf(ListViewGroups.lsvgUI))
            NewTabWriterItem("Office Symbol:", user.GetAttribute("physicalDeliveryOfficeName"), NameOf(ListViewGroups.lsvgUI))
        End If
    End Sub

    Private Sub ExportBasicInfo(sender As Object, e As DoWorkEventArgs)
        Dim userSelectedCSVFilePath As String = e.Argument
        MyBase.ExportFromListView(Me.BasicInfoListView, userSelectedCSVFilePath)
    End Sub

End Class