﻿'------------------------------------------------------------------------------
' <auto-generated>
'     This code was generated by a tool.
'     Runtime Version:4.0.30319.42000
'
'     Changes to this file may cause incorrect behavior and will be lost if
'     the code is regenerated.
' </auto-generated>
'------------------------------------------------------------------------------

Option Strict On
Option Explicit On


Namespace My
    
    <Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute(),  _
     Global.System.CodeDom.Compiler.GeneratedCodeAttribute("Microsoft.VisualStudio.Editors.SettingsDesigner.SettingsSingleFileGenerator", "16.8.1.0"),  _
     Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)>  _
    Partial Friend NotInheritable Class MySettings
        Inherits Global.System.Configuration.ApplicationSettingsBase
        
        Private Shared defaultInstance As MySettings = CType(Global.System.Configuration.ApplicationSettingsBase.Synchronized(New MySettings()),MySettings)
        
#Region "My.Settings Auto-Save Functionality"
#If _MyType = "WindowsForms" Then
    Private Shared addedHandler As Boolean

    Private Shared addedHandlerLockObject As New Object

    <Global.System.Diagnostics.DebuggerNonUserCodeAttribute(), Global.System.ComponentModel.EditorBrowsableAttribute(Global.System.ComponentModel.EditorBrowsableState.Advanced)> _
    Private Shared Sub AutoSaveSettings(sender As Global.System.Object, e As Global.System.EventArgs)
        If My.Application.SaveMySettingsOnExit Then
            My.Settings.Save()
        End If
    End Sub
#End If
#End Region
        
        Public Shared ReadOnly Property [Default]() As MySettings
            Get
                
#If _MyType = "WindowsForms" Then
               If Not addedHandler Then
                    SyncLock addedHandlerLockObject
                        If Not addedHandler Then
                            AddHandler My.Application.Shutdown, AddressOf AutoSaveSettings
                            addedHandler = True
                        End If
                    End SyncLock
                End If
#End If
                Return defaultInstance
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("C:\ProgramData\chocolatey\lib\psexec\tools\PsExec.exe")>  _
        Public Property PsExecPath() As String
            Get
                Return CType(Me("PsExecPath"),String)
            End Get
            Set
                Me("PsExecPath") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("LDAP://DC=HBC,DC=local")>  _
        Public ReadOnly Property LDAPEntryPath() As String
            Get
                Return CType(Me("LDAPEntryPath"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property RecentTasks() As Global.System.Collections.Specialized.StringCollection
            Get
                Return CType(Me("RecentTasks"),Global.System.Collections.Specialized.StringCollection)
            End Get
            Set
                Me("RecentTasks") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property CustomActions() As Global.System.Collections.Specialized.StringCollection
            Get
                Return CType(Me("CustomActions"),Global.System.Collections.Specialized.StringCollection)
            End Get
            Set
                Me("CustomActions") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\hbcfs1\HBCShared\ITTeam\logs")>  _
        Public ReadOnly Property LogsRemotePath() As String
            Get
                Return CType(Me("LogsRemotePath"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("IT\Logs")>  _
        Public ReadOnly Property LogsLocalPathWithoutSystemDrive() As String
            Get
                Return CType(Me("LogsLocalPathWithoutSystemDrive"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("IT")>  _
        Public ReadOnly Property ITFolderLocalPathWithoutSystemDrive() As String
            Get
                Return CType(Me("ITFolderLocalPathWithoutSystemDrive"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Software\SysTool3\Actions")>  _
        Public ReadOnly Property RegistryPathCustomActions() As String
            Get
                Return CType(Me("RegistryPathCustomActions"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Software\SysTool3\Collections")>  _
        Public ReadOnly Property RegistryPathCollections() As String
            Get
                Return CType(Me("RegistryPathCollections"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Software\SysTool3\Queries")>  _
        Public ReadOnly Property RegistryPathQueries() As String
            Get
                Return CType(Me("RegistryPathQueries"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("Software\SysTool3")>  _
        Public ReadOnly Property RegistryPath() As String
            Get
                Return CType(Me("RegistryPath"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("630")>  _
        Public Property MainFormHeight() As Integer
            Get
                Return CType(Me("MainFormHeight"),Integer)
            End Get
            Set
                Me("MainFormHeight") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("1030")>  _
        Public Property MainFormWidth() As Integer
            Get
                Return CType(Me("MainFormWidth"),Integer)
            End Get
            Set
                Me("MainFormWidth") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("HBC.local")>  _
        Public ReadOnly Property DomainName() As String
            Get
                Return CType(Me("DomainName"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property ActiveComputers() As Global.System.Collections.Specialized.StringCollection
            Get
                Return CType(Me("ActiveComputers"),Global.System.Collections.Specialized.StringCollection)
            End Get
            Set
                Me("ActiveComputers") = value
            End Set
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute()>  _
        Public Property ActiveCollections() As Global.System.Collections.Specialized.StringCollection
            Get
                Return CType(Me("ActiveCollections"),Global.System.Collections.Specialized.StringCollection)
            End Get
            Set
                Me("ActiveCollections") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("\\hbcfs1\HBCShared\ITTeam\Software\SysTool\upgrade.ps1")>  _
        Public ReadOnly Property UpgradeScriptPath() As String
            Get
                Return CType(Me("UpgradeScriptPath"),String)
            End Get
        End Property
        
        <Global.System.Configuration.UserScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("True")>  _
        Public Property UserSettingsMigrateRequired() As Boolean
            Get
                Return CType(Me("UserSettingsMigrateRequired"),Boolean)
            End Get
            Set
                Me("UserSettingsMigrateRequired") = value
            End Set
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("http://tfs/City%20of%20Davenport/_packaging/9e1bb9d1-5fa5-4d46-93fc-b747c423d783/"& _ 
            "nuget/v3/")>  _
        Public ReadOnly Property NuGetURI() As String
            Get
                Return CType(Me("NuGetURI"),String)
            End Get
        End Property
        
        <Global.System.Configuration.ApplicationScopedSettingAttribute(),  _
         Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
         Global.System.Configuration.DefaultSettingValueAttribute("it-systemnfos3")>  _
        Public ReadOnly Property NuGetPkgName() As String
            Get
                Return CType(Me("NuGetPkgName"),String)
            End Get
        End Property
    End Class
End Namespace

Namespace My
    
    <Global.Microsoft.VisualBasic.HideModuleNameAttribute(),  _
     Global.System.Diagnostics.DebuggerNonUserCodeAttribute(),  _
     Global.System.Runtime.CompilerServices.CompilerGeneratedAttribute()>  _
    Friend Module MySettingsProperty
        
        <Global.System.ComponentModel.Design.HelpKeywordAttribute("My.Settings")>  _
        Friend ReadOnly Property Settings() As Global.systemnfos3.My.MySettings
            Get
                Return Global.systemnfos3.My.MySettings.Default
            End Get
        End Property
    End Module
End Namespace
