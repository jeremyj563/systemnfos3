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
    Public Property ActiveDirectoryPath As String
    Public Property ActiveDirectoryContainer As LDAPContainer
    Public Property ConnectionString As String
    Private Property LDAPData As SearchResult

    Public Sub New(ldapData As SearchResult)
        Me.LDAPData = ldapData

        Dim netAddrs = Me.GetValue(Me.LDAPData.Properties("networkAddress")).Split(",")
        Dim desc = Me.GetValue(Me.LDAPData.Properties("description"))
        Me.IPAddress = netAddrs(0)
        Me.MACAddress = If(netAddrs.Count > 1, netAddrs(1), String.Empty)
        Me.Value = Me.GetValue(Me.LDAPData.Properties("name"))
        Me.Description = If(String.IsNullOrWhiteSpace(desc), "Unknown", desc)
        Me.Display = $"{Me.Description}  >  {Me.Value}"
        Me.UserName = Me.GetValue(Me.LDAPData.Properties("uid"))
        Me.DisplayName = Me.GetValue(Me.LDAPData.Properties("displayName"))
        Me.LastLogon = Me.GetValue(Me.LDAPData.Properties("extensionAttribute1"))
        Me.ActiveDirectoryPath = If(Me.LDAPData.Path, String.Empty)
        Me.ActiveDirectoryContainer = New LDAPContainer(Me.LDAPData.Path)
        Me.ConnectionString = Me.Value
    End Sub

    Private Function GetValue(values As ResultPropertyValueCollection) As String
        If values?.Count() > 0 Then
            Dim isArray = values(0).GetType().IsArray AndAlso values(0).Count() > 0
            Return If(isArray, values(0)(0), values(0))
        End If
        Return String.Empty
    End Function

End Class