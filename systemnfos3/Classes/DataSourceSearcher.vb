Imports System.ComponentModel
Imports System.Net.NetworkInformation
Imports System.Text.RegularExpressions

Public Class DataSourceSearcher

    Private Property SearchTerm As String = Nothing
    Private Property BindingList As New BindingList(Of DataUnit)()

    Public Sub New(searchTerm As String, bindingSource As BindingSource)
        Me.SearchTerm = searchTerm.Trim()

        ' Must create a copy of the bound data since the collection could change when the 
        ' asynchronous LDAPMonitor code runs in MainForm.UpdateLDAPDataBindings()
        Me.BindingList.AddRange(bindingSource.DataSource)
    End Sub

    Public Function GetComputers() As List(Of Computer)
        ' This function returns a list of computer(s) found by searching through the data source using the provided search term
        Dim searchResults As New List(Of Computer)()

        If Me.SearchTerm.Trim().Equals("/") OrElse Me.SearchTerm.Trim().Equals("\") Then
            ' Search term implied "root" so display all computers
            searchResults.AddRange(Me.BindingList.OfType(Of Computer)())
        Else
            ' Use a regular expression to check if the search term is a *full* IPv4 address (all four octects present)
            ' Matching pattern taken from "Regular Expressions Cookbook 2nd Edition" by Jan Goyvaerts & Steven Levithan - Page 469 Chapter 8.16: Matching IPv4 Addresses
            If Regex.Matches(Me.SearchTerm, "^(?:(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)\.){3}(?:25[0-5]|2[0-4][0-9]|[01]?[0-9][0-9]?)$").Count = 1 Then
                ' Search term was an IPv4 address so ping it
                Dim ping As New Ping()
                If ping.Send(Me.SearchTerm, 500).Status = IPStatus.Success Then
                    Try
                        ' Device reponded to ping so attempt to connect via WMI and lookup the ComputerName
                        Dim wmi As New WMIController(Me.SearchTerm, WMIController.ManagementScopes.Regular, async:=True)
                        Dim wmiComputerName As String = wmi.GetPropertyValue(wmi.Query("SELECT Name FROM Win32_ComputerSystem"), "Name")
                        If Not String.IsNullOrWhiteSpace(wmiComputerName) Then
                            ' WMI connection successful so set the search term to the ComputerName that was retrieved
                            Me.SearchTerm = wmiComputerName
                        End If
                    Catch ex As Exception
                        ' WMI connection failed so notify the user and clear the search term
                        MsgBox("Responded but could not connect.", icon:=MessageBoxIcon.Exclamation, caption:=Me.SearchTerm)
                        Me.SearchTerm = String.Empty
                    End Try
                Else
                    ' Ping failed so attempt to match the IPv4 search term to the 'networkAddress' LDAP attribute that was loaded at launch
                    Dim matchingComputer As Computer = Me.BindingList.OfType(Of Computer)().SingleOrDefault(Function(c) c.IPAddress = Me.SearchTerm)
                    If matchingComputer IsNot Nothing Then
                        ' The 'networkAddress' LDAP attribute matched so set the search term to the ComputerName value of the matching computer object
                        Me.SearchTerm = matchingComputer.Value
                    Else
                        ' No match found so notify the user and clear the search term
                        MsgBox("Did not respond and no offline data found.", icon:=MessageBoxIcon.Error, caption:=Me.SearchTerm)
                        Me.SearchTerm = String.Empty
                    End If
                End If
            End If

            ' If the search term is now empty then simply return empty search results
            If Not String.IsNullOrWhiteSpace(Me.SearchTerm) Then
                ' Check if the search term is an exact computer name match
                Dim matchingComputer = Me.BindingList.OfType(Of Computer)().SingleOrDefault(Function(c) c.Value.ToUpper().Equals(Me.SearchTerm.ToUpper()))
                If matchingComputer IsNot Nothing Then
                    ' The search term matches *exactly* one computer in the data source
                    Try
                        ' Check if the computer responds to WMI and that the search term matches the ComputerName returned by WMI
                        Dim wmi As New WMIController(Me.SearchTerm, WMIController.ManagementScopes.Regular, async:=True)
                        Dim wmiComputerName As String = wmi.GetPropertyValue(wmi.Query("SELECT Name FROM Win32_ComputerSystem"), "Name")
                        If Not String.IsNullOrWhiteSpace(wmiComputerName) Then
                            If wmiComputerName = matchingComputer.Value Then
                                ' The ComputerName returned by WMI was a match so add the matching computer object to the search results
                                searchResults.Add(matchingComputer)
                            Else
                                ' The ComputerName returned by WMI was NOT a match so don't add it to the search results and instead alert the user
                                Dim message = "Responded but the requested computer name does not match the resonse from WMI!"
                                MsgBox(message, icon:=MessageBoxIcon.Error, caption:=$"Requested: {Me.SearchTerm} Response: {wmiComputerName}")
                            End If
                        End If
                    Catch ex As Exception
                        ' The WMI connection failed but still add the matching computer object to the search results so we can load its "offline" LDAP data
                        searchResults.Add(matchingComputer)
                    End Try
                Else
                    ' No exact match found so get a list of computer(s) that contain the search term
                    searchResults.AddRange(Me.BindingList.OfType(Of Computer)().Where(Function(c) _
                        c.Value.ToUpper().Contains(Me.SearchTerm.ToUpper()) OrElse
                        c.Description.ToUpper().Contains(Me.SearchTerm.ToUpper()) OrElse
                        c.Display.ToUpper().Contains(Me.SearchTerm.ToUpper()) OrElse
                        c.DisplayName.ToUpper().Contains(Me.SearchTerm.ToUpper()) OrElse
                        c.IPAddress.ToUpper().Contains(Me.SearchTerm.ToUpper()) OrElse
                        c.UserName.ToUpper().Contains(Me.SearchTerm.ToUpper())))
                End If
            End If
        End If

        Return searchResults
    End Function

    Public Function GetComputer() As Computer
        For Each computer As Computer In Me.BindingList.OfType(Of Computer)()
            Dim computerName As String = computer.Value.ToUpper().Trim()
            If computerName = Me.SearchTerm.ToUpper().Trim() Then
                Return computer
            End If
        Next

        Return Nothing
    End Function

End Class