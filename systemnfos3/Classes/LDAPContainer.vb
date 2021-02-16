Imports System.Reflection

Public Class LDAPContainer

    Public Property Description
        Get
            Return GetAttribute("description")
        End Get
        Set(value)
            SetAttribute("description", value)
        End Set
    End Property

    Public ReadOnly Property LastLogonExtensionAttribute
        Get
            Return GetAttribute("extensionAttribute1")
        End Get
    End Property

    Public Property PhysicalDeliveryOfficeName
        Get
            Return GetAttribute("physicalDeliveryOfficeName")
        End Get
        Set(value)
            SetAttribute("physicalDeliveryOfficeName", value)
        End Set
    End Property

    Private Property DirectoryEntry As DirectoryEntry

    Public Sub New(activeDirectoryObject As DirectoryEntry)
        Me.DirectoryEntry = New DirectoryEntry(activeDirectoryObject)
    End Sub

    ''' <summary>
    ''' If you do not have a DirectoryEntry or an ADsPath loaded, one can be built. The format of searching AD for your item will be (SearchTerm=SearchValue)
    ''' </summary>
    ''' <param name="objectClass">The type of AD object you are looking for. This will be in the beginning of the query</param>
    ''' <param name="searchTerm">The attribute you will be searching against</param>
    ''' <param name="searchValue">The value of the attribute you are seaerching against</param>
    ''' <remarks></remarks>
    Public Sub New(objectClass As String, searchTerm As String, searchValue As String)
        Me.New($"(&(objectClass={objectClass})({searchTerm}={searchValue}))")
    End Sub

    Public Sub New(queryOrPath As String)
        If queryOrPath.StartsWith("LDAP") Then
            ' Used for loading ldap data at launch and during incremental updates
            Me.DirectoryEntry = New DirectoryEntry(queryOrPath)
        Else
            ' Used for querying data of a specific ldap object
            Using entry As New DirectoryEntry() With {.AuthenticationType = AuthenticationTypes.Secure}
                Using searcher As New DirectorySearcher(entry) With {.Filter = queryOrPath}
                    Dim searchResult As SearchResult = searcher.FindOne()
                    If searchResult IsNot Nothing Then
                        Me.DirectoryEntry = New DirectoryEntry(searchResult.Path)
                    End If
                End Using
            End Using
        End If
    End Sub

    Public Sub SetAttribute(attributeName As String, value As Object)
        Try
            Me.DirectoryEntry.Properties(attributeName).Value = value
            Me.DirectoryEntry.CommitChanges()
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Function GetAttribute(attributeName As String) As Object
        Try
            If Me.DirectoryEntry IsNot Nothing Then
                Dim response As Object = Me.DirectoryEntry.Properties(attributeName).Value
                If response IsNot Nothing Then
                    If response.GetType().Name.ToUpper() = "__COMOBJECT" Then
                        Using searcher As New DirectorySearcher(DirectoryEntry)
                            searcher.PropertiesToLoad.Add(attributeName)

                            Return searcher.FindOne().Properties(attributeName)(0)
                        End Using
                    Else
                        Return response
                    End If
                End If
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return Nothing
    End Function

End Class
