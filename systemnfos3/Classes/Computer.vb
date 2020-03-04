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
        SetProperties()
    End Sub

    Private Sub SetProperties()
        SetNetworkAddressProperties()

        Dim description = GetResultPropertyValue(Me.LDAPData.Properties("description"))

        Me.Value = GetResultPropertyValue(Me.LDAPData.Properties("name"))
        Me.Description = If(String.IsNullOrEmpty(description), "Unknown", description)
        Me.Display = String.Format("{0}  >  {1}", Me.Description, Me.Value)
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
            Me.IPAddress = networkAddresses(0)
            If networkAddresses.Count = 2 Then
                Me.MACAddress = networkAddresses(1)
            End If
        End If

        If Me.IPAddress Is Nothing Then Me.IPAddress = String.Empty
        If Me.MACAddress Is Nothing Then Me.MACAddress = String.Empty
    End Sub

    Private Function GetResultPropertyValue(properties As ResultPropertyValueCollection) As String
        Dim retVal As String = String.Empty

        If properties IsNot Nothing Then
            If properties.Count > 0 Then
                If properties(0).GetType().IsArray Then
                    retVal = CType(properties(0), String())(0)
                Else
                    retVal = CType(properties(0), String)
                End If
            End If
        End If

        Return retVal
    End Function

End Class