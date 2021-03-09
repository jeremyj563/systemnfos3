Public Module LDAPSearcher

    Private Function NewBaseDirectorySearcher(Optional filter As String = "") As DirectorySearcher
        Dim baseSearcher As New DirectorySearcher()
        With baseSearcher
            .Filter = $"(&(objectClass=computer){filter})"
            .PropertiesToLoad.AddRange({"name", "description", "uid", "displayName", "extensionAttribute1", "networkAddress", "whenChanged"})
            .PageSize = 1000
            .SizeLimit = 5000
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
            Using searcher As DirectorySearcher = NewBaseDirectorySearcher()

                Return searcher.FindAllSortBy("description")
            End Using
        End Using
    End Function

    Public Function GetOneComputerResult(computerName As String) As SearchResult
        Using entry As DirectoryEntry = NewBaseDirectoryEntry()
            Dim filter = $"&(name={computerName})"
            Using searcher As DirectorySearcher = NewBaseDirectorySearcher(filter)

                Return searcher.FindOne()
            End Using
        End Using
    End Function

    Public Function GetIncrementalUpdateResults(timestamp As String) As List(Of SearchResult)
        Using entry As DirectoryEntry = NewBaseDirectoryEntry()
            Dim convertedTimestamp = ManagementDateTimeConverter.ToDmtfDateTime(Convert.ToDateTime(timestamp)).Split(".")(0)
            Dim filter = $"&(!whenChanged<={convertedTimestamp}.0Z)"
            Using searcher As DirectorySearcher = NewBaseDirectorySearcher(filter)

                Return searcher.FindAllSortBy("description")
            End Using
        End Using
    End Function

    Public Function GetLastChangedTime(results As List(Of SearchResult)) As Long
        Dim maxADDateTime As Long
        Dim i As Integer = 0
        For Each result As SearchResult In results
            Dim Value, Display, SN, Username, DisplayName, MAC, IP, WhenChanged As String


            Value = result.Properties("name").Item(0)
            If result.Properties("description").Count > 0 Then
                DisplayName = result.Properties("description").Item(0)
            Else
                DisplayName = "Unknown"
            End If
            If Trim(DisplayName) = "" Then DisplayName = "Unknown"
            Display = DisplayName & "  >  " & Value
            If result.Properties("carLicense").Count > 0 Then
                SN = result.Properties("carLicense").Item(0)
            Else
                SN = Nothing
            End If
            If result.Properties("uid").Count > 0 Then
                Username = result.Properties("uid").Item(0)
            Else
                Username = Nothing
            End If
            If result.Properties("networkAddress").Count > 0 Then
                Dim splNA As String() = result.Properties("networkAddress").Item(0).ToString.Split(",")
                If splNA.Count > 1 Then
                    IP = Trim(splNA(0))
                    MAC = Trim(splNA(1))
                ElseIf splNA.Count = 1 Then
                    IP = Nothing
                    MAC = result.Properties("networkAddress").Item(0)
                Else
                    IP = Nothing
                    MAC = Nothing
                End If
            Else
                IP = Nothing
                MAC = Nothing
            End If
            If result.Properties("whenChanged").Count > 0 Then
                Dim Holder As DateTime = result.Properties("whenChanged").Item(0)
                Dim Month As String
                Dim Day As String
                Dim Hour As String
                Dim Minute As String
                Dim Second As String

                If Holder.Month.ToString.Length = 1 Then
                    Month = 0 & Holder.Month
                Else
                    Month = Holder.Month
                End If
                If Holder.Day.ToString.Length = 1 Then
                    Day = 0 & Holder.Day
                Else
                    Day = Holder.Day
                End If
                If Holder.Hour.ToString.Length = 1 Then
                    Hour = 0 & Holder.Hour
                Else
                    Hour = Holder.Hour
                End If
                If Holder.Minute.ToString.Length = 1 Then
                    Minute = 0 & Holder.Minute
                Else
                    Minute = Holder.Minute
                End If
                If Holder.Second.ToString.Length = 1 Then
                    Second = 0 & Holder.Second
                Else
                    Second = Holder.Second
                End If
                Holder = result.Properties("whenChanged").Item(0)
                WhenChanged = Holder.Year & Month & Day & Hour & Minute & Second
                If maxADDateTime < CType(WhenChanged, Long) Then maxADDateTime = CType(WhenChanged, Long)
            Else
                WhenChanged = Nothing
            End If
            i += 1
        Next

        Return maxADDateTime
    End Function

End Module