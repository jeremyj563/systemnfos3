Imports Microsoft.Win32
Imports System.Reflection

Public Class RegistryController

    Private Property Scope As ManagementScope

    Public Sub New()
        Dim options As New ConnectionOptions With
            {
                .Impersonation = ImpersonationLevel.Impersonate,
                .Authentication = AuthenticationLevel.Packet,
                .EnablePrivileges = True,
                .Timeout = New TimeSpan(0, 0, 0, 5, 0)
            }

        Me.Scope = New ManagementScope("\\.\root\default")
        With Me.Scope
            .Options = options
            .Options.Context.Add("__ProviderArchitecture", 32)
            .Connect()
        End With
    End Sub

    Public Sub New(scope As ManagementScope)
        Me.Scope = scope
    End Sub

    Public Sub New(computerName As String, Optional architecture As RegistryArchitectures = RegistryArchitectures.x86)
        Dim options As New ConnectionOptions With
            {
                .Impersonation = ImpersonationLevel.Impersonate,
                .Authentication = AuthenticationLevel.Packet,
                .EnablePrivileges = True,
                .Timeout = New TimeSpan(0, 0, 0, 5, 0)
            }

        Me.Scope = New ManagementScope($"\\{computerName}\root\default")
        With Me.Scope
            .Options = options
            .Options.Context.Add("__ProviderArchitecture", architecture)
            .Connect()
        End With
    End Sub

    Public Function GetKeyValues(keyPath As String, Optional hive As RegistryHive = RegistryHive.LocalMachine, Optional methodName As String = "EnumKey") As String()
        Dim keyValues As String() = New String() {}

        Try
            Dim baseProperties As Dictionary(Of String, Object) = NewBaseWMIProperties(hive, keyPath)
            Dim wmi As WMIManagement = WMIManagementFactory(methodName, baseProperties)
            Dim baseObject As ManagementBaseObject = wmi.ManagementClass.InvokeMethod(methodName, wmi.ManagementBaseObject, wmi.InvokeMethodOptions)

            Dim properties As String() = baseObject.Properties("sNames").Value
            If properties IsNot Nothing Then
                keyValues = properties.Select(Function(p) p.ToUpper().Trim()).ToArray()
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return keyValues
    End Function

    Public Function GetKeyValue(keyPath As String, valueName As String, Optional valueType As RegistryKeyValueTypes = RegistryKeyValueTypes.Dword, Optional hive As RegistryHive = RegistryHive.LocalMachine) As Object
        Dim methodName As String = FindMethodName(valueType, verb:="Get")

        Try
            Dim properties As Dictionary(Of String, Object) = NewBaseWMIProperties(hive, keyPath)
            properties.Add("sValueName", valueName)
            Dim wmi = WMIManagementFactory(methodName, properties)
            wmi.ManagementBaseObject = wmi.ManagementClass.InvokeMethod(methodName, wmi.ManagementBaseObject, wmi.InvokeMethodOptions)

            If valueType = RegistryKeyValueTypes.Dword Then
                Return wmi.ManagementBaseObject("uValue")
            End If

            If valueType = RegistryKeyValueTypes.String Then
                Return wmi.ManagementBaseObject("sValue")
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return Nothing
    End Function

    Public Sub SetKeyValue(keyPath As String, valueName As String, value As String, Optional valueType As RegistryKeyValueTypes = RegistryKeyValueTypes.Dword, Optional hive As RegistryHive = RegistryHive.LocalMachine)
        Dim methodName As String = FindMethodName(valueType, verb:="Set")

        Try
            Dim properties As Dictionary(Of String, Object) = NewBaseWMIProperties(hive, keyPath)
            properties.Add("sValueName", valueName)

            If valueType = RegistryKeyValueTypes.Dword Then
                properties.Add("uValue", value)
            End If

            If valueType = RegistryKeyValueTypes.String Then
                properties.Add("sValue", value)
            End If

            Dim wmi = WMIManagementFactory(methodName, properties)
            wmi.ManagementClass.InvokeMethod(methodName, wmi.ManagementBaseObject, wmi.InvokeMethodOptions)
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Sub NewKey(keyName As String, Optional hive As RegistryHive = RegistryHive.LocalMachine)
        Dim methodName As String = "CreateKey"

        Try
            Dim properties As Dictionary(Of String, Object) = NewBaseWMIProperties(hive, keyName)

            Dim wmi = WMIManagementFactory(methodName, properties)
            wmi.ManagementClass.InvokeMethod(methodName, wmi.ManagementBaseObject, wmi.InvokeMethodOptions)
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Sub DeleteKey(keyPath As String, Optional hive As RegistryHive = RegistryHive.LocalMachine)
        Dim methodName As String = "DeleteKey"

        Try
            Dim keyValues As String() = GetKeyValues(keyPath, hive)
            For Each key As String In keyValues
                DeleteKey($"{keyPath}\{key}", hive)
            Next

            Dim properties As Dictionary(Of String, Object) = NewBaseWMIProperties(hive, keyPath)

            Dim wmi = WMIManagementFactory(methodName, properties)
            wmi.ManagementClass.InvokeMethod(methodName, wmi.ManagementBaseObject, wmi.InvokeMethodOptions)
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Function FindMethodName(valueType As RegistryKeyValueTypes, verb As String) As String
        ' Return a method name for getting/setting either a DWORD or String value
        Select Case valueType
            Case RegistryKeyValueTypes.Dword
                Return $"{verb}DWORDValue"
            Case RegistryKeyValueTypes.String
                Return $"{verb}StringValue"
        End Select

        Return Nothing
    End Function

    Private Function NewBaseWMIProperties(hive As RegistryHive, keyPath As String) As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object)() From
        {
            {"hDefKey", Convert.ToInt64(hive.ToString("X"), 16)}, ' .ToString("X") converts to Hexadecimal value
            {"sSubKeyName", keyPath}
        }
    End Function

    Private Function WMIManagementFactory(methodName As String, properties As Dictionary(Of String, Object)) As WMIManagement
        Dim objectGetOptions = New ObjectGetOptions()
        objectGetOptions.Context.Add("__ProviderArchitecture", Me.Scope.Options.Context("__ProviderArchitecture"))

        Dim invokeMethodOptions = New InvokeMethodOptions()
        invokeMethodOptions.Context.Add("__ProviderArchitecture", Me.Scope.Options.Context("__ProviderArchitecture"))

        Dim managementClass = New ManagementClass("stdRegProv", objectGetOptions)
        managementClass.Scope = Me.Scope

        Dim managementBaseObject As ManagementBaseObject = managementClass.GetMethodParameters(methodName)
        For Each [property] As KeyValuePair(Of String, Object) In properties
            managementBaseObject.SetPropertyValue([property].Key, [property].Value)
        Next

        Dim wmiManagement As New WMIManagement() With
        {
            .ManagementClass = managementClass,
            .ManagementBaseObject = managementBaseObject,
            .InvokeMethodOptions = invokeMethodOptions
        }

        Return wmiManagement
    End Function

    Public Enum RegistryArchitectures
        x86 = 32
        x64 = 64
    End Enum

    Public Enum RegistryKeyValueTypes
        [String] = 1
        Dword = 2
    End Enum

    Public Enum RegistrySettings
        UpdateAvailable
    End Enum

End Class