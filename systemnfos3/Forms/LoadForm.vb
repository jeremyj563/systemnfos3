Imports System.Threading
Imports System.Reflection
Imports System.ComponentModel
Imports Microsoft.Management.Infrastructure
Imports Microsoft.Win32
Imports System.Text.RegularExpressions
Imports System.Threading.Tasks

Public NotInheritable Class LoadForm

    Private Property DataSource As New BindingList(Of DataUnit)()
    Private Property LastChangedLDAPTime As Date
    Private Property MouseDrag As Boolean
    Private Property MouseX As Integer
    Private Property MouseY As Integer

#Region " Form "

    Private Async Sub Me_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.Text = "SysTool SplashScreen"

        ' Prepare global logger
        [Global].LogMaintenance([Global].LogPath)
        [Global].SetLog4NetFileAppenderPath([Global].LogPath)

        ' Create base registry keys if they don't already exist
        NewAppRegistryKey(RegistryHive.LocalMachine)
        NewAppRegistryKey(RegistryHive.CurrentUser)

        ' Execute the application upgrade script if an update is available
        'UpgradeApp()

        ' Authenticate the current user
        'AuthenticateCurrentUser()

        ' Migrate user settings from previous version
        'MigrateUserSettings()

        ' User is authorized so begin loading data into the data source
        Await Task.Run(Sub() GetAllData())
        NewMainForm()
    End Sub

    Private Sub NewMainForm()
        FadeOut()

        Dim mainForm As New MainForm(Me.DataSource, Me.LastChangedLDAPTime)
        mainForm.Show()

        Me.Close()
        Me.Dispose()
    End Sub

#End Region

#Region " Application "

    Private Sub NewAppRegistryKey(hive As RegistryHive)
        Dim registry As New RegistryController()
        Dim regPath As String() = My.Settings.RegistryPath.Split("\")

        Dim baseKey = regPath.First().ToUpper()
        Dim keyValues = registry.GetKeyValues(baseKey, hive)

        Dim app = regPath.Last().ToUpper()
        If Not keyValues.Any(Function(value) value = app) Then
            ' Base registry key doesn't exist so create a new one
            registry.NewKey(My.Settings.RegistryPath, hive)
        End If
    End Sub

    Private Sub UpgradeApp()
        Try
            Dim registry As New RegistryController()
            Dim regEntry As String = NameOf(RegistryController.RegistrySettings.UpdateAvailable)
            Dim updateAvailable As Boolean = registry.GetKeyValue(My.Settings.RegistryPath, regEntry, hive:=RegistryHive.LocalMachine)

            If updateAvailable Then
                Process.Start("powershell.exe", "-ExecutionPolicy ByPass -File " & My.Settings.UpgradeScriptPath)

                ' Set the registry entry to *not* trigger upgrade upon next launch
                registry.SetKeyValue(My.Settings.RegistryPath, regEntry, "0", hive:=RegistryHive.LocalMachine)

                ' Upgrade script was successfully executed so exit the program
                ExitApp(0, "Upgrading application...", [Global].LogEvents.Info)
            End If
        Catch ex As Exception
            ' Error encountered running upgrade script so log the error and continue running the program
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
            Exit Sub
        End Try
    End Sub

    Private Sub MigrateUserSettings()
        If My.Settings.UserSettingsMigrateRequired Then
            With My.Settings
                .Upgrade()
                .UserSettingsMigrateRequired = False
                .Save()
            End With
        End If
    End Sub

#End Region

#Region " Authentication "

    Private Sub AuthenticateCurrentUser()
        Dim authorizedDomain As String = "DC=HBC,DC=local"
        Dim authorizedGroupPath As String = "OU=HBC_UsersOU"
        Dim authorizedGroupName As String = "SysTool"

        If Not UserInAuthorizedGroup(Environment.UserName, authorizedDomain, authorizedGroupPath, authorizedGroupName) Then
            Dim message = $"Your user account ({Environment.UserName}) is not authorized to access this program. Please contact the IT department for assistance."
            MessageBox.Show(Me, message, "Unauthorized", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
            [Global].ExitApp(99, "Unauthorized User", [Global].LogEvents.Fatal)
        End If
    End Sub

    Private Function UserInAuthorizedGroup(userName As String, domainName As String, groupPath As String, groupName As String) As Boolean
        ' Lookup the user's ldap group membership via local WMI provider
        Try
            Dim session = CimSession.Create("localhost")
            Dim query = $"SELECT DS_memberOf FROM ds_user WHERE ds_samAccountName='{userName}'"
            Dim user = session.QueryInstances("root\directory\ldap", "WQL", query).SingleOrDefault()

            If user IsNot Nothing Then
                Dim dnGroups As String() = user.CimInstanceProperties("DS_memberOf").Value

                ' RegEx distinguishedName matching pattern taken from https://regexr.com/3l4au
                Dim pattern = "^(?:(?<cn>CN=(?<name>[^,]*)),)?(?:(?<path>(?:(?:CN|OU)=[^,]+,?)+),)?(?<domain>(?:DC=[^,]+,?)+)$"

                For Each dn As String In dnGroups
                    Dim matches = Regex.Matches(dn, pattern).Cast(Of Match)()

                    Dim match As Match = matches.SingleOrDefault(Function(m) _
                        m.Groups("domain").Value = domainName AndAlso
                        m.Groups("path").Value = groupPath AndAlso
                        m.Groups("name").Value = groupName)

                    If match IsNot Nothing Then
                        Return True
                    End If
                Next
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return False
    End Function

#End Region

#Region " Data Source "

    Private Sub GetAllData()
        GetLDAPComputers()
        GetRegistryCollections()
        GetRegistryQueries()
        GetRegistryCustomActions()
    End Sub

    Private Sub GetLDAPComputers()
        Using searchResults = LDAPSearcher.GetAllComputerResults()
            If searchResults IsNot Nothing Then
                Me.LastChangedLDAPTime = LDAPSearcher.GetLastChangedTime(searchResults)

                For Each result As SearchResult In searchResults
                    Me.DataSource.Add(New Computer(result))
                Next
            End If
        End Using
    End Sub

    Private Sub GetRegistryCollections()
        Try
            Dim registry As New RegistryController()
            Dim keyValues = registry.GetKeyValues(My.Settings.RegistryPath, RegistryHive.CurrentUser)

            Dim pathCollections As String = My.Settings.RegistryPathCollections.Split("\").Last()
            If Not keyValues.Any(Function(value) value = pathCollections) Then
                ' Collections registry key doesn't exist so create it
                registry.NewKey(My.Settings.RegistryPathCollections, RegistryHive.CurrentUser)
            Else
                ' Collections registry key DOES exist so load any stored collections
                Dim collectionKeyValues As String() = registry.GetKeyValues(My.Settings.RegistryPathCollections, RegistryHive.CurrentUser)
                For Each collectionName As String In collectionKeyValues
                    Me.DataSource.Add(New Collection() With
                        {
                            .Value = collectionName,
                            .Display = $"COLL:{collectionName}"
                        })
                Next
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Sub GetRegistryQueries()
        Dim registry As New RegistryController()
        Dim keyValues = registry.GetKeyValues(My.Settings.RegistryPath, RegistryHive.CurrentUser)

        Dim pathQueries As String = My.Settings.RegistryPathQueries.Split("\").Last()
        If Not keyValues.Any(Function(key) key = pathQueries) Then
            ' Queries registry key doesn't exist so create it
            registry.NewKey(My.Settings.RegistryPathQueries, RegistryHive.CurrentUser)
        End If
    End Sub

    Private Sub GetRegistryCustomActions()
        Dim registry As New RegistryController()
        Dim keyValues = registry.GetKeyValues(My.Settings.RegistryPath, RegistryHive.CurrentUser)

        Dim pathCustomActions As String = My.Settings.RegistryPathCustomActions.Split("\").Last()
        If Not keyValues.Any(Function(key) key = pathCustomActions) Then
            ' Custom actions registry key doesn't exist so create it
            registry.NewKey(My.Settings.RegistryPathCustomActions, RegistryHive.CurrentUser)
        End If
    End Sub

#End Region

#Region " Splash Screen Effects "

    Private Sub FadeOut()
        For value = 100 To 0 Step -2
            SetOpacity(value)
            Thread.Sleep(10)
        Next
    End Sub

    Private Sub SetOpacity(value As Integer)
        Me.Opacity = value / 100
    End Sub

    Private Sub Me_MouseDown(sender As Object, e As MouseEventArgs) Handles Me.MouseDown
        MouseDrag = True
        MouseX = Cursor.Position.X - Me.Left
        MouseY = Cursor.Position.Y - Me.Top
    End Sub

    Private Sub Me_MouseMove(sender As Object, e As MouseEventArgs) Handles Me.MouseMove
        If MouseDrag Then
            Me.Top = Cursor.Position.Y - MouseY
            Me.Left = Cursor.Position.X - MouseX
        End If
    End Sub

    Private Sub Me_MouseUp(sender As Object, e As MouseEventArgs) Handles Me.MouseUp
        MouseDrag = False
    End Sub

#End Region

End Class