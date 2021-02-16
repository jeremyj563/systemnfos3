Imports System.Data.Entity.Design.PluralizationServices
Imports System.Environment
Imports System.Globalization
Imports System.IO
Imports System.Reflection
Imports log4net
Imports log4net.Appender
Imports log4net.Config
Imports log4net.Repository.Hierarchy

Public Module [Global]

    Public Property AppPath As String = GetAppPath()
    Public Property AppDir As String = GetAppDir()
    Public Property AppVersion As String = GetAppVersion()
    Public Property AppName As String = GetAppName()
    Public Property LogPath As String = GetLogPath()
    Public Property SystemDrive As String = GetSystemDrive()
    Public Property ITFolderLocalPath As String = GetITFolderPath()
    Public Property MainForm As Form

#Region " Logging "

    Public Sub LogMaintenance(logPath As String)
        Dim archivePath As String = Path.Combine(Path.GetDirectoryName(logPath), "archive")

        Try
            ' Create archive path if it doesn't exist
            If Not Directory.Exists(archivePath) Then
                Directory.CreateDirectory(archivePath)
            End If

            ' Do not continue if current log isn't found
            If File.Exists(logPath) Then
                ' Timestamp and move the current log if it's 10 days or older
                If (Date.Now - File.GetCreationTime(logPath)).TotalDays >= 10 Then
                    Dim outFileName = $"{Path.GetFileNameWithoutExtension(logPath)}.{Date.Now:MMddyyyyHHmmss}{Path.GetExtension(logPath)}"
                    Dim outPath = Path.Combine(archivePath, outFileName)

                    File.Move(logPath, outPath)
                End If

                ' Purge any log files that are 1 year or older
                For Each file As String In Directory.GetFiles(archivePath)
                    If (Date.Now - IO.File.GetCreationTime(file)).TotalDays >= 365 Then
                        IO.File.Delete(file)
                    End If
                Next
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Private Function GetLogPath() As String
        ' Log to remote path if available, otherwise log to local path
        Dim logsFolderPath As String = Nothing

        Try
            If Directory.Exists(My.Settings.LogsRemotePath) AndAlso UserHasWriteAccessToPath(My.Settings.LogsRemotePath) Then
                logsFolderPath = Path.Combine(My.Settings.LogsRemotePath, [Global].AppName)
            Else
                Dim logsLocalPath As String = Path.Combine(Path.GetPathRoot(Environment.SystemDirectory), My.Settings.LogsLocalPathWithoutSystemDrive)
                If Not String.IsNullOrWhiteSpace(logsLocalPath) AndAlso Directory.Exists(logsLocalPath) Then
                    logsFolderPath = Path.Combine(logsLocalPath, [Global].AppName)
                Else
                    ExitApp(1, errorWritingToLog:=True)
                End If
            End If
        Catch ex As Exception
            ExitApp(1, ex.Message, errorWritingToLog:=True)
        End Try

        Return Path.Combine(logsFolderPath, $"{[Global].AppName}.log")
    End Function

    Public Sub SetLog4NetFileAppenderPath(logPath As String)
        XmlConfigurator.Configure()

        Dim appender As FileAppender = CType(LogManager.GetRepository(), Hierarchy).Root.Appenders _
            .ToArray() _
            .Where(Function(a) TypeOf a Is FileAppender) _
            .First()

        appender.File = logPath
        appender.ActivateOptions()
    End Sub

    Public Sub LogEvent(message As String, Optional logEvent As LogEvents = LogEvents.Error)
        message = $"{message} | {[Global].AppVersion} | {Environment.UserName}@{Environment.MachineName}"
        Dim logger As ILog = LogManager.GetLogger("FileAppender")

        Try
            Select Case logEvent
                Case LogEvents.Debug
                    logger.DebugFormat(message)
                Case LogEvents.Error
                    logger.ErrorFormat(message)
                Case LogEvents.Fatal
                    logger.FatalFormat(message)
                Case LogEvents.Info
                    logger.InfoFormat(message)
                Case LogEvents.Warn
                    logger.WarnFormat(message)
            End Select

        Catch ex As Exception
            ExitApp(1, errorWritingToLog:=True)
        End Try
    End Sub

    Public Enum LogEvents
        Debug
        [Error]
        Fatal
        Info
        Warn
    End Enum

#End Region

#Region " Static Methods "

    Public Function MsgBox(message As String, Optional buttons As MessageBoxButtons = MessageBoxButtons.OK, Optional icon As MessageBoxIcon = MessageBoxIcon.Information, Optional caption As String = Nothing) As DialogResult
        If caption Is Nothing Then
            caption = Application.ProductName
        End If

        Dim result As DialogResult = Nothing
        [Global].MainForm.UIThread(Sub() result = MessageBox.Show([Global].MainForm, message, caption, buttons, icon))

        Return result
    End Function

    Public Function UserHasWriteAccessToPath(path As String) As Boolean
        Try
            Directory.GetAccessControl(path)
        Catch ex As Exception
            Return False
        End Try

        Return True
    End Function

    Public Function Pluralize(text As String) As String
        Dim pluralizedText As String = Nothing

        Dim cultureInfo As New CultureInfo("en-us")
        Dim PluralizationService As PluralizationService = PluralizationService.CreateService(cultureInfo)
        If PluralizationService.IsSingular(text) Then
            pluralizedText = PluralizationService.Pluralize(text)
        End If

        Return pluralizedText
    End Function

    Public Enum Nodes
        Computers
        Collections
        Task
        Query
        Results
        RootNode
        CustomActions
        Settings
    End Enum

    Public Enum Switch
        NewInstance
    End Enum

    Public Sub SetPsExecBinaryPath()
        Try
            If String.IsNullOrWhiteSpace(My.Settings.PsExecPath) OrElse Not File.Exists(My.Settings.PsExecPath) Then
                Dim possiblePsExecLocations As New List(Of String) From
                {
                    Path.Combine(SpecialFolder.CommonApplicationData, "chocolatey\lib\psexec\tools\PsExec.exe"),
                    Path.Combine(Environment.SystemDirectory, "PsExec.exe"),
                    Path.Combine([Global].ITFolderLocalPath, "PsExec.exe"),
                    Path.Combine([Global].AppDir, "PsExec.exe")
                }

                For Each location As String In possiblePsExecLocations
                    If File.Exists(location) Then
                        ' Found an already existing copy of PsExec.exe
                        My.Settings.PsExecPath = location
                        My.Settings.Save()
                        Exit For
                    End If
                Next

                If String.IsNullOrWhiteSpace(My.Settings.PsExecPath) OrElse Not File.Exists(My.Settings.PsExecPath) Then
                    ' Did not find any local copy of PsExec.exe so prompt the user to select the file location
                    Dim openDialog As New OpenFileDialog() With
                    {
                        .Filter = "PSEXEC.EXE|PsExec.exe",
                        .Title = "Provide PsExec.exe location",
                        .CheckFileExists = True,
                        .Multiselect = False
                    }

                    If openDialog.ShowDialog() = DialogResult.OK Then
                        If openDialog.FileName.ToUpper().EndsWith("PSEXEC.EXE") Then
                            My.Settings.PsExecPath = openDialog.FileName
                            My.Settings.Save()
                        End If
                    Else
                        Exit Sub
                    End If
                End If
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Sub WriteListOfObjectsToCSV(Of T)(items As IEnumerable(Of T), path As String, Optional includeHeader As Boolean = False)
        Try
            Dim properties = GetType(T).GetProperties(BindingFlags.[Public] Or BindingFlags.Instance)

            Using writer As New StreamWriter(path)
                If includeHeader Then
                    writer.WriteLine(String.Join(",", properties.Select(Function([Property]) [Property].Name)))
                End If

                For Each item In items
                    writer.WriteLine(String.Join(",", properties.Select(Function([Property]) [Property].GetValue(item, Nothing))))
                Next
            End Using

        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Sub WriteDataTableToCSV(table As DataTable, path As String, Optional includeHeader As Boolean = False)
        Try
            Using writer As New StreamWriter(path)
                If includeHeader Then
                    Dim headers = table.Columns.Cast(Of DataColumn)() _
                        .Select(Function(column) column.ColumnName) _
                        .ToArray()

                    writer.WriteLine(String.Join(",", headers))
                End If

                For Each row As DataRow In table.Rows
                    Dim items = row.ItemArray _
                        .Select(Function(item) If(item.ToString().Contains(","), $"'{item}'", item)) _
                        .ToArray()

                    writer.WriteLine(String.Join(",", items))
                Next
            End Using
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try
    End Sub

    Public Sub SetControlDoubleBufferedProperty(control As Control, Optional value As Boolean = True)
        control.GetType().InvokeMember("DoubleBuffered", BindingFlags.SetProperty Or BindingFlags.Instance Or BindingFlags.NonPublic, Nothing, control, New Object() {value})
    End Sub

    Public Sub ExitApp(exitCode As Integer, Optional message As String = "", Optional logEvent As LogEvents = LogEvents.Error, Optional errorWritingToLog As Boolean = False)
        If exitCode = 0 Then
            message = "completed successfully"
        End If

        message = $"[{[Global].AppName}] {message} - exit code: {exitCode}"
        If errorWritingToLog Then
            Console.WriteLine(message)
        Else
            [Global].LogEvent(message, logEvent)
        End If

        Environment.Exit(exitCode)
    End Sub

#End Region

#Region " Private Methods "

    Private Function GetAppPath() As String
        Dim appPath As String = Nothing

        Try
            appPath = Process.GetCurrentProcess().MainModule.FileName
        Catch ex As Exception
            ExitApp(1, $"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return appPath
    End Function

    Private Function GetAppDir() As String
        Dim appDir As String = Nothing

        Try
            appDir = Path.GetDirectoryName(AppPath)
        Catch ex As Exception
            ExitApp(1, $"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return appDir
    End Function

    Private Function GetAppVersion() As String
        Dim version As String = "not found"

        Try
            Dim nuspecFile As String = Directory _
                .GetFiles([Global].AppDir, "*.nuspec") _
                .FirstOrDefault()

            If Not String.IsNullOrWhiteSpace(nuspecFile) Then
                Dim xmlText As String = File.ReadAllText(nuspecFile)
                Dim xmlElements As XElement = XElement.Parse(xmlText)
                version = CType(xmlElements.Elements.Nodes(1), XElement).Value.Trim()
            End If
        Catch ex As Exception
            LogEvent($"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return version
    End Function

    Private Function GetAppName() As String
        Dim appName As String = Nothing

        Try
            Dim assemblyPath = Assembly.GetEntryAssembly().Location
            appName = Path.GetFileNameWithoutExtension(assemblyPath)
        Catch ex As Exception
            ExitApp(1, $"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return appName
    End Function

    Private Function GetSystemDrive() As String
        Dim systemDrive As String = Nothing

        Try
            systemDrive = Directory.GetDirectoryRoot(Environment.SystemDirectory)
        Catch ex As Exception
            ExitApp(1, $"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return systemDrive
    End Function

    Private Function GetITFolderPath() As String
        Dim itPath As String = Nothing

        Try
            itPath = Path.Combine(SystemDrive, My.Settings.ITFolderLocalPathWithoutSystemDrive)
        Catch ex As Exception
            ExitApp(1, $"EXCEPTION in {MethodBase.GetCurrentMethod()}: {ex.Message}")
        End Try

        Return itPath
    End Function

#End Region

End Module