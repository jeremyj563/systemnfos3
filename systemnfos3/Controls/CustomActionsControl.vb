Imports systemnfos3.RegistryController
Imports Microsoft.Win32
Imports System.IO

Public Class CustomActionsControl

    Public Property OwnerForm As MainForm = Nothing
    Public Property OwnerNode As TreeNode = Nothing

    Private Property ChangeExisting As Boolean = False
    Private Property NewCustomActionName As String = "New Custom Action"
    Private Property OldCustomActionName As String = Nothing

    Public Sub New(ownerForm As MainForm)
        ' This call is required by the designer.
        InitializeComponent()

        Me.OwnerForm = ownerForm
    End Sub

    Public Sub New(ownerForm As MainForm, existingCustomActionName As String, ownerNode As TreeNode)
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.ChangeExisting = True
        Me.NewCustomActionName = existingCustomActionName
        Me.OwnerForm = ownerForm
        Me.OwnerNode = ownerNode
        Me.OldCustomActionName = existingCustomActionName
    End Sub

    Private Sub CustomActionNameTextBox_TextChanged(sender As Object, e As EventArgs) Handles CustomActionNameTextBox.TextChanged
        Me.CustomActionNameLabel.Text = Me.CustomActionNameTextBox.Text
        Me.NewCustomActionName = Me.CustomActionNameTextBox.Text
    End Sub

    Private Sub SaveButton_Click(sender As Object, e As EventArgs) Handles SaveButton.Click
        Dim registry As New RegistryController()
        Dim customActions As String() = registry.GetKeyValues(My.Settings.RegistryPathCustomActions, RegistryHive.CurrentUser)

        Dim oldCustomActionKey As String = Path.Combine(My.Settings.RegistryPathCustomActions, OldCustomActionName)
        Dim newCustomActionKey As String = Path.Combine(My.Settings.RegistryPathCustomActions, NewCustomActionName)

        If Me.ChangeExisting Then
            If customActions.Any(Function(customAction) customAction.ToUpper() = Me.OldCustomActionName.ToUpper()) Then
                registry.DeleteKey(oldCustomActionKey, RegistryHive.CurrentUser)
            End If

            registry.NewKey(newCustomActionKey, RegistryHive.CurrentUser)
        Else
            If customActions.Any(Function(customAction) customAction.ToUpper() = Me.NewCustomActionName.ToUpper()) Then
                MsgBox("Custom action name already exists. Please choose a different name, or delete the old custom action")
                Exit Sub
            Else
                registry.NewKey(newCustomActionKey, RegistryHive.CurrentUser)
            End If
        End If

        If Not String.IsNullOrWhiteSpace(Me.TechCommandsTextBox.Text) Then
            Dim i As Integer = 1
            For Each customAction As String In Me.TechCommandsTextBox.Text.Split(Environment.NewLine)
                registry.SetKeyValue(newCustomActionKey, i, customAction, RegistryKeyValueTypes.String, RegistryHive.CurrentUser)
                i += 1
            Next
        End If

        If Not Me.ChangeExisting Then
            Dim newTreeNode As New TreeNode()
            Me.OwnerNode = newTreeNode
            Me.ChangeExisting = True

            Dim customActionOptionsMenuStrip As New ContextMenuStrip()
            customActionOptionsMenuStrip.Items.Add("Delete", Nothing, Sub() DeleteCustomAction(registry, newTreeNode, newCustomActionKey))

            With newTreeNode
                .Name = NewCustomActionName
                .Text = NewCustomActionName
                .Tag = Me
                .ContextMenuStrip = customActionOptionsMenuStrip
            End With

            Me.OwnerForm.ResourceExplorer.Nodes(NameOf(Nodes.Settings)).Nodes(NameOf(Nodes.CustomActions)).Nodes.Add(newTreeNode)
        Else
            With Me.OwnerNode
                .Tag = Me
                .Text = NewCustomActionName
                .Name = NewCustomActionName
            End With
        End If

    End Sub

    Private Sub DeleteCustomAction(ByRef registryContext As RegistryController, treeNode As TreeNode, customActionKey As String)
        registryContext.DeleteKey(customActionKey, RegistryHive.CurrentUser)
        treeNode.Remove()
    End Sub

    Private Sub ControlCustomActions_Load(sender As Object, e As EventArgs) Handles Me.Load
        Me.CustomActionNameTextBox.Text = Me.NewCustomActionName
        Me.CustomActionNameLabel.Text = Me.NewCustomActionName

        If Me.ChangeExisting Then
            Dim registry As New RegistryController()
            Dim customActionExists As Boolean = False

            For Each customAction As String In registry.GetKeyValues(My.Settings.RegistryPathCustomActions, RegistryHive.CurrentUser)
                If NewCustomActionName.ToUpper() = customAction.ToUpper() Then
                    customActionExists = True
                    Exit For
                End If
            Next

            If customActionExists Then
                Dim textOutput As String = Nothing
                Dim newCustomActionKey As String = Path.Combine(My.Settings.RegistryPathCustomActions, NewCustomActionName)

                Dim values As String() = registry.GetKeyValues(newCustomActionKey, RegistryHive.CurrentUser, methodName:="EnumValues")
                For i = 0 To values.Count() - 1
                    textOutput += registry.GetKeyValue(newCustomActionKey, values(i), RegistryKeyValueTypes.String, RegistryHive.CurrentUser)
                    If i > values.Count() - 1 Then textOutput += Environment.NewLine
                Next

                Me.TechCommandsTextBox.Text = textOutput
            Else
                MsgBox("Could not find custom action!")
                Me.OwnerForm.ResourceExplorer.Nodes(NameOf(Nodes.Settings)).Nodes(NameOf(Nodes.CustomActions)).Nodes(NewCustomActionName).Remove()
            End If
        End If
    End Sub

    Private Sub CustomActionsControl_VisibleChanged(sender As Object, e As EventArgs) Handles Me.VisibleChanged
        If Me.Visible Then
            Me.OwnerForm.AcceptButton = Nothing
        End If

        If Not Me.Visible Then
            Me.OwnerForm.AcceptButton = Me.OwnerForm.SubmitButton
        End If
    End Sub

End Class