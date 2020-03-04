<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CustomActionsControl
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.CustomActionNameTextBox = New System.Windows.Forms.TextBox()
        Me.CustomActionNameLabel = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.TechCommandsTextBox = New System.Windows.Forms.TextBox()
        Me.SaveButton = New System.Windows.Forms.Button()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'TextBoxCustomActionName
        '
        Me.CustomActionNameTextBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CustomActionNameTextBox.Location = New System.Drawing.Point(322, 54)
        Me.CustomActionNameTextBox.Name = "TextBoxCustomActionName"
        Me.CustomActionNameTextBox.Size = New System.Drawing.Size(108, 20)
        Me.CustomActionNameTextBox.TabIndex = 0
        Me.CustomActionNameTextBox.Text = "New Custom Action"
        '
        'LabelCustomActionName
        '
        Me.CustomActionNameLabel.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.CustomActionNameLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 15.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.CustomActionNameLabel.Location = New System.Drawing.Point(8, 0)
        Me.CustomActionNameLabel.Name = "LabelCustomActionName"
        Me.CustomActionNameLabel.Size = New System.Drawing.Size(732, 25)
        Me.CustomActionNameLabel.TabIndex = 1
        Me.CustomActionNameLabel.Text = "New Custom Action"
        Me.CustomActionNameLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(319, 38)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(106, 13)
        Me.Label1.TabIndex = 2
        Me.Label1.Text = "Custom Action Name"
        '
        'TextBoxTechCommands
        '
        Me.TechCommandsTextBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TechCommandsTextBox.Location = New System.Drawing.Point(3, 117)
        Me.TechCommandsTextBox.Name = "TextBoxTechCommands"
        Me.TechCommandsTextBox.Size = New System.Drawing.Size(742, 20)
        Me.TechCommandsTextBox.TabIndex = 4
        '
        'ButtonSave
        '
        Me.SaveButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SaveButton.Location = New System.Drawing.Point(670, 143)
        Me.SaveButton.Name = "ButtonSave"
        Me.SaveButton.Size = New System.Drawing.Size(75, 23)
        Me.SaveButton.TabIndex = 5
        Me.SaveButton.Text = "Save"
        Me.SaveButton.UseVisualStyleBackColor = True
        '
        'Label2
        '
        Me.Label2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(3, 101)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(250, 13)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Type in command to PSEXEC on a remote computer"
        '
        'ControlCustomActions
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.SaveButton)
        Me.Controls.Add(Me.TechCommandsTextBox)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.CustomActionNameLabel)
        Me.Controls.Add(Me.CustomActionNameTextBox)
        Me.Name = "ControlCustomActions"
        Me.Size = New System.Drawing.Size(748, 169)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents CustomActionNameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents CustomActionNameLabel As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents TechCommandsTextBox As System.Windows.Forms.TextBox
    Friend WithEvents SaveButton As System.Windows.Forms.Button
    Friend WithEvents Label2 As System.Windows.Forms.Label

End Class
