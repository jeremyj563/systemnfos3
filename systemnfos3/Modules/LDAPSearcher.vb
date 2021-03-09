Public Module LDAPSearcher

    Private Function NewBaseDirectorySearcher(filter As String) As DirectorySearcher
        Dim baseSearcher As New DirectorySearcher()
        With baseSearcher
            .Filter = $"(&(objectClass=computer){filter})"
            .PropertiesToLoad.AddRange({"name", "description", "uid", "displayName", "extensionAttribute1", "networkAddress", "whenChanged"})
            .PageSize = 1000
        End With

        Return baseSearcher
    End Function

    Private Function NewBaseDirectoryEntry() As DirectoryEntry
        Dim baseEntry As New DirectoryEntry(My.Settings.LDAPEntryPath)
        baseEntry.AuthenticationType = AuthenticationTypes.FastBind Or AuthenticationTypes.Secure

        Return baseEntry
    End Function

    Public Function GetAllComputerResults() As List(Of SearchResult)
        Using entry As DirectoryEntry = NewBaseDirectoryEntry()
            Using searcher As DirectorySearcher = NewBaseDirectorySearcher(My.Settings.LDAPComputerFilter)

                Return searcher.FindAllSortBy("description")
            End Using
        End Using
    End Function

    Public Function GetOneComputerResult(computerName As String) As SearchResult
        Using entry As DirectoryEntry = NewBaseDirectoryEntry()
            Dim filter = $"(&(objectClass=computer)(&(name={computerName})))"
            Using searcher As DirectorySearcher = NewBaseDirectorySearcher(filter)

                Return searcher.FindOne()
            End Using
        End Using
    End Function

    Public Function GetIncrementalUpdateResults(timestamp As Date) As List(Of SearchResult)
        Using entry As DirectoryEntry = NewBaseDirectoryEntry()
            Using searcher As DirectorySearcher = NewBaseDirectorySearcher(My.Settings.LDAPComputerFilter)
                Dim convertedTimestamp = ManagementDateTimeConverter.ToDmtfDateTime(timestamp).Split(".")(0)
                searcher.Filter = $"(&(objectClass=computer)(!whenChanged<={convertedTimestamp}.0Z){My.Settings.LDAPComputerFilter})"

                Return searcher.FindAllSortBy("description")
            End Using
        End Using
    End Function

    Public Function GetLastChangedTime(results As List(Of SearchResult)) As Date
        ' This date/time value is used to query ldap for incremental updates in MainForm.UpdateLDAPDataBindings()
        Dim lastChangedResult = results.OfType(Of SearchResult)() _
            .Where(Function(searchResult) searchResult.Properties("whenChanged").Count > 0) _
            .OrderByDescending(Function(searchResult) CType(searchResult.Properties("whenChanged")(0), Date)) _
            .First()

        Dim lastChangedTime As Date = lastChangedResult.Properties("whenChanged")(0)

        Return lastChangedTime
    End Function

End Module