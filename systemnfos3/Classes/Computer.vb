Public Class Computer
    Inherits DataUnit

    Public Overrides Property Value As String
    Public Overrides Property Display As String
    Public Property UserName As String
    Public Property DisplayName As String
    Public Property Description As String
    Public Property LastLogon As String
    Public Property MACAddress As String
    Public Property IPAddress As String
    Public Property LDAPContainer As LDAPContainer
    Public Property ConnectionString As String

    Public Sub New(ldapData As SearchResult)
        Dim netAddrs = Me.GetValue(ldapData.Properties("networkAddress")).Split(",")
        Dim desc = Me.GetValue(ldapData.Properties("description"))
        Me.IPAddress = netAddrs(0)
        Me.MACAddress = If(netAddrs.Count > 1, netAddrs(1), String.Empty)
        Me.Value = Me.GetValue(ldapData.Properties("name"))
        Me.Description = If(String.IsNullOrWhiteSpace(desc), "Unknown", desc)
        Me.Display = $"{Me.Description}  >  {Me.Value}"
        Me.UserName = Me.GetValue(ldapData.Properties("uid"))
        Me.DisplayName = Me.GetValue(ldapData.Properties("displayName"))
        Me.LastLogon = Me.GetValue(ldapData.Properties("extensionAttribute1"))
        Me.LDAPContainer = New LDAPContainer(ldapData.Path)
        Me.ConnectionString = Me.Value
    End Sub

    Public Sub New(value As String, display As String, userName As String, displayName As String, macAddress As String, ipAddress As String, ldapData As SearchResult)
        Me.Value = value
        Me.Display = display
        Me.UserName = userName
        Me.DisplayName = displayName
        Me.MACAddress = macAddress
        Me.IPAddress = ipAddress
        Me.LDAPContainer = New LDAPContainer(ldapData.Path)
    End Sub

    Private Function GetValue(values As ResultPropertyValueCollection) As String
        If values?.Count() > 0 Then
            Dim isArray = values(0).GetType().IsArray AndAlso values(0).Count() > 0
            Return If(isArray, values(0)(0), values(0))
        End If
        Return String.Empty
    End Function

End Class