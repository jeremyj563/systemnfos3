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
        Me.SetProperties()
    End Sub

    Private Sub SetProperties()
        Me.SetNetworkAddressProperties()

        Dim description = GetResultPropertyValue(Me.LDAPData.Properties("description"))
        Me.Value = GetResultPropertyValue(Me.LDAPData.Properties("name"))
        Me.Description = If(String.IsNullOrWhiteSpace(description), "Unknown", description)
        Me.Display = $"{Me.Description}  >  {Me.Value}"
        Me.UserName = GetResultPropertyValue(Me.LDAPData.Properties("uid"))
        Me.DisplayName = GetResultPropertyValue(Me.LDAPData.Properties("displayName"))
        Me.LastLogon = GetResultPropertyValue(Me.LDAPData.Properties("extensionAttribute1"))
        Me.ActiveDirectoryPath = If(Me.LDAPData.Path, String.Empty)
        Me.ActiveDirectoryContainer = New LDAPContainer(Me.LDAPData.Path)
        Me.ConnectionString = Me.Value
    End Sub

    Private Sub SetNetworkAddressProperties()
        Dim networkAddress As String = GetResultPropertyValue(Me.LDAPData.Properties("networkAddress"))
        If Not String.IsNullOrWhiteSpace(networkAddress) Then
            Dim networkAddresses As String() = networkAddress.Split(",")
            Me.IPAddress = If(networkAddresses(0), String.Empty)
            If networkAddresses.Count >= 2 Then
                Me.MACAddress = If(networkAddresses(1), String.Empty)
            End If
        End If
    End Sub

    Private Function GetResultPropertyValue(properties As ResultPropertyValueCollection) As String
        If properties IsNot Nothing AndAlso properties.Count > 0 Then
            If properties(0).GetType().IsArray Then
                Return CType(properties(0), String())(0)
            Else
                Return CType(properties(0), String)
            End If
        End If

        Return String.Empty
    End Function

End Class