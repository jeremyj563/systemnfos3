<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskControl
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
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.ArgumentsLabel = New System.Windows.Forms.Label()
        Me.FileComboBox = New System.Windows.Forms.ComboBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.ArgumentsTextBox = New System.Windows.Forms.TextBox()
        Me.chbViewable = New System.Windows.Forms.CheckBox()
        Me.ArgumentsCheckBox = New System.Windows.Forms.CheckBox()
        Me.BrowseButton = New System.Windows.Forms.Button()
        Me.btnStart = New System.Windows.Forms.Button()
        Me.CollectionsTreeView = New System.Windows.Forms.TreeView()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.TaskNameTextBox = New System.Windows.Forms.TextBox()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.TaskNameTextBox)
        Me.GroupBox1.Controls.Add(Me.Label3)
        Me.GroupBox1.Controls.Add(Me.Label2)
        Me.GroupBox1.Controls.Add(Me.ArgumentsLabel)
        Me.GroupBox1.Controls.Add(Me.FileComboBox)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.ArgumentsTextBox)
        Me.GroupBox1.Controls.Add(Me.chbViewable)
        Me.GroupBox1.Controls.Add(Me.ArgumentsCheckBox)
        Me.GroupBox1.Controls.Add(Me.BrowseButton)
        Me.GroupBox1.Location = New System.Drawing.Point(207, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(516, 193)
        Me.GroupBox1.TabIndex = 26
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Program/Command Parameters"
        '
        'lblArguments
        '
        Me.ArgumentsLabel.AutoSize = True
        Me.ArgumentsLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ArgumentsLabel.Location = New System.Drawing.Point(112, 134)
        Me.ArgumentsLabel.Name = "lblArguments"
        Me.ArgumentsLabel.Size = New System.Drawing.Size(298, 12)
        Me.ArgumentsLabel.TabIndex = 7
        Me.ArgumentsLabel.Text = "- Check this box to add argument parameters to your program/command"
        '
        'cbxFile
        '
        Me.FileComboBox.AutoCompleteMode = System.Windows.Forms.AutoCompleteMode.SuggestAppend
        Me.FileComboBox.AutoCompleteSource = System.Windows.Forms.AutoCompleteSource.ListItems
        Me.FileComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FileComboBox.FormattingEnabled = True
        Me.FileComboBox.Location = New System.Drawing.Point(6, 89)
        Me.FileComboBox.Name = "cbxFile"
        Me.FileComboBox.Size = New System.Drawing.Size(440, 21)
        Me.FileComboBox.TabIndex = 6
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(112, 174)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(303, 12)
        Me.Label1.TabIndex = 5
        Me.Label1.Text = "- Check this box to make the operation viewable on the remote system(s)"
        '
        'txtArguments
        '
        Me.ArgumentsTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ArgumentsTextBox.Location = New System.Drawing.Point(103, 130)
        Me.ArgumentsTextBox.Name = "txtArguments"
        Me.ArgumentsTextBox.Size = New System.Drawing.Size(344, 20)
        Me.ArgumentsTextBox.TabIndex = 4
        Me.ArgumentsTextBox.Visible = False
        '
        'chbViewable
        '
        Me.chbViewable.AutoSize = True
        Me.chbViewable.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.chbViewable.Location = New System.Drawing.Point(6, 172)
        Me.chbViewable.Name = "chbViewable"
        Me.chbViewable.Size = New System.Drawing.Size(99, 17)
        Me.chbViewable.TabIndex = 3
        Me.chbViewable.Text = "Make Viewable"
        Me.chbViewable.UseVisualStyleBackColor = True
        '
        'chbArguments
        '
        Me.ArgumentsCheckBox.AutoSize = True
        Me.ArgumentsCheckBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ArgumentsCheckBox.Location = New System.Drawing.Point(6, 132)
        Me.ArgumentsCheckBox.Name = "chbArguments"
        Me.ArgumentsCheckBox.Size = New System.Drawing.Size(76, 17)
        Me.ArgumentsCheckBox.TabIndex = 2
        Me.ArgumentsCheckBox.Text = "Arguments"
        Me.ArgumentsCheckBox.UseVisualStyleBackColor = True
        '
        'btnBrowse
        '
        Me.BrowseButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.BrowseButton.Location = New System.Drawing.Point(452, 89)
        Me.BrowseButton.Name = "btnBrowse"
        Me.BrowseButton.Size = New System.Drawing.Size(55, 23)
        Me.BrowseButton.TabIndex = 1
        Me.BrowseButton.Text = "Browse"
        Me.BrowseButton.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.Location = New System.Drawing.Point(207, 202)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(71, 23)
        Me.btnStart.TabIndex = 27
        Me.btnStart.Text = "Start Task"
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'tvwCollections
        '
        Me.CollectionsTreeView.Location = New System.Drawing.Point(3, 3)
        Me.CollectionsTreeView.Name = "tvwCollections"
        Me.CollectionsTreeView.Size = New System.Drawing.Size(198, 518)
        Me.CollectionsTreeView.TabIndex = 28
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 6.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 60)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(504, 26)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "If running a file such as an executable, batch or printerExport (cab) then the fi" & _
    "le must be selected from a remote location. Ex: \\server\folder\file.exe"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(6, 22)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(65, 13)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "Task Name:"
        '
        'txtTaskName
        '
        Me.TaskNameTextBox.Location = New System.Drawing.Point(78, 19)
        Me.TaskNameTextBox.Name = "txtTaskName"
        Me.TaskNameTextBox.Size = New System.Drawing.Size(369, 20)
        Me.TaskNameTextBox.TabIndex = 10
        '
        'msTaskControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.CollectionsTreeView)
        Me.Controls.Add(Me.btnStart)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "msTaskControl"
        Me.Size = New System.Drawing.Size(726, 524)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents ArgumentsLabel As System.Windows.Forms.Label
    Friend WithEvents FileComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents ArgumentsTextBox As System.Windows.Forms.TextBox
    Friend WithEvents chbViewable As System.Windows.Forms.CheckBox
    Friend WithEvents ArgumentsCheckBox As System.Windows.Forms.CheckBox
    Friend WithEvents BrowseButton As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents CollectionsTreeView As System.Windows.Forms.TreeView
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents TaskNameTextBox As System.Windows.Forms.TextBox

End Class
