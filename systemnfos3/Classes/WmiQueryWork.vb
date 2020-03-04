Imports System.Reflection

Public Class WMIQueryWork

    Private Property WMI As WMIController

    Public Function GetWMIResult(serviceName As String, computerName As String) As String
        WMI = New WMIController(computerName)
        Try
            Select Case serviceName
                Case "NAC IP Address"
                    Return WMI.GetPropertyValue(WMI.Query(String.Format("SELECT IPAddress FROM Win32_NetworkAdapterConfiguration WHERE DNSDomain='{0}'", My.Settings.DomainName)), "IPAddress")
                Case "NAC MAC Address"
                    Return WMI.GetPropertyValue(WMI.Query(String.Format("SELECT MACAddress FROM Win32_NetworkAdapterConfiguration WHERE DNSDomain='{0}'", My.Settings.DomainName)), "MACAddress")
                Case "NAC Default Gateway"
                    Return WMI.GetPropertyValue(WMI.Query(String.Format("SELECT DefaultIPGateway FROM Win32_NetworkAdapterConfiguration WHERE DNSDomain='{0}'", My.Settings.DomainName)), "DefaultIPGateway")
                Case "OS Name"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT Caption FROM Win32_OperatingSystem"), "Caption")
                Case "OS Description"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT Description FROM Win32_OperatingSystem"), "Description")
                Case "OS Architecture"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT AddressWidth FROM Win32_Processor"), "AddressWidth")
                Case "CS Bios Version"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT SMBIOSBIOSVERSION FROM Win32_BIOS"), "SMBIOSBIOSVersion")
                Case "CS Model"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT Model FROM Win32_ComputerSystem"), "Model")
                Case "CS Serial Number"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT SerialNumber FROM Win32_BIOS"), "SerialNumber")
                Case "CS Image Date"
                    Return DateConverter(WMI.GetPropertyValue(WMI.Query("SELECT InstallDate FROM Win32_OperatingSystem"), "InstallDate"))
                Case "CS Last Boot"
                    Return DateConverter(WMI.GetPropertyValue(WMI.Query("SELECT LastBootupTime From Win32_OperatingSystem"), "LastBootupTime"))
                Case "CS Total HD Space"
                    Return Convert.ToInt32(WMI.GetPropertyValue(WMI.Query("SELECT Capacity FROM Win32_Volume WHERE BootVolume=True"), "Capacity") / 1073741824) & " GBs"
                Case "CS Free HD Space"
                    Return Convert.ToInt32(WMI.GetPropertyValue(WMI.Query("Select FreeSpace FROM Win32_Volume WHERE BootVolume=True"), "FreeSpace") / 1073741824) & " GBs"
                Case "CS CPU"
                    Return WMI.GetPropertyValue(WMI.Query("SELECT Name FROM Win32_Processor"), "Name")
                Case "CS Total Memory"
                    Return Convert.ToInt32(WMI.GetPropertyValue(WMI.Query("SELECT Capacity FROM Win32_PhysicalMemory"), "Capacity") / 1073741824) & " GBs"
                Case "CS Bitlocker Status"
                    Return GetBitLockerInfo()
            End Select

        Catch ex As Exception
            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
        End Try

        Return Nothing
    End Function

    Private Function DateConverter(originalText As String, Optional includeTime As Boolean = False) As String
        Dim fullDate As String = Nothing
        Dim year As String
        Dim month As String
        Dim day As String

        If originalText Is Nothing Then
            Return Nothing
        Else
            year = CType(Mid(originalText, 1, 4), Integer)
            month = CType(Mid(originalText, 5, 2), Integer)
            day = CType(Mid(originalText, 7, 2), Integer)
            fullDate = String.Format("{0}/{1}/{2}", month, day, year)
        End If

        If includeTime Then
            fullDate += String.Format(" {0}:{1}:{2}", Mid(originalText, 9, 2), Mid(originalText, 11, 2), Mid(originalText, 13, 2))
        End If

        Return fullDate
    End Function

    Private Function GetBitLockerInfo() As String
        Dim bitLocker As ManagementClass = New ManagementClass(WMI.BitLockerScope, WMI.BitLockerScope.Path, New ObjectGetOptions)
        If WMI.BitLockerScope.IsConnected Then
            If bitLocker.GetInstances.Count > 0 Then
                Dim bitLockerEnabled As Boolean = False
                For Each bitLockerInstance As ManagementObject In bitLocker.GetInstances()
                    Dim protectionStatus(0) As Object
                    bitLockerInstance.InvokeMethod("GetProtectionStatus", protectionStatus)

                    Dim hddEncryptionStatus(1) As Object
                    bitLockerInstance.InvokeMethod("GetConversionStatus", hddEncryptionStatus)

                    If protectionStatus(0) = 1 Then
                        bitLockerEnabled = True
                        Select Case hddEncryptionStatus(0)
                            Case 0
                                Return "Enabled - Fully Decrypted"
                            Case 1
                                Return "Enabled - Fully Encrypted"
                            Case 2
                                Return String.Format("Enabled - Encrypting at {0}%", hddEncryptionStatus(1))
                            Case 3
                                Return String.Format("Enabled - Decrypting at {0}%", hddEncryptionStatus(1))
                            Case 4
                                Return String.Format("Enabled - Encrypting paused at {0}%", hddEncryptionStatus(1))
                            Case 5
                                Return String.Format("Enabled - Decrypting paused at {0}%", hddEncryptionStatus(1))
                        End Select
                        Exit For
                    Else

                    End If
                Next
                If Not bitLockerEnabled Then Return "Disabled"
            End If
        End If

        Return Nothing
    End Function

End Class