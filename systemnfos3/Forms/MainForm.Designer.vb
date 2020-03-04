<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class MainForm
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
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
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim TreeNode1 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Computers")
        Dim TreeNode2 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Collections")
        Dim TreeNode3 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Task")
        Dim TreeNode4 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Query")
        Dim TreeNode5 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Results")
        Dim TreeNode6 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Resource Explorer", New System.Windows.Forms.TreeNode() {TreeNode1, TreeNode2, TreeNode3, TreeNode4, TreeNode5})
        Dim TreeNode7 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Custom Actions")
        Dim TreeNode8 As System.Windows.Forms.TreeNode = New System.Windows.Forms.TreeNode("Settings", New System.Windows.Forms.TreeNode() {TreeNode7})
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(MainForm))
        Me.MainSplitContainer = New System.Windows.Forms.SplitContainer()
        Me.ResourceExplorer = New System.Windows.Forms.TreeView()
        Me.UserInputComboBox = New System.Windows.Forms.ComboBox()
        Me.LabelResource = New System.Windows.Forms.Label()
        Me.SubmitButton = New System.Windows.Forms.Button()
        Me.ClearButton = New System.Windows.Forms.Button()
        Me.NewButton = New System.Windows.Forms.Button()
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip()
        Me.UpTimeStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.MemoryUsageStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.CpuUsageStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.ConnectionsStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        Me.DateTimeStatusLabel = New System.Windows.Forms.ToolStripStatusLabel()
        CType(Me.MainSplitContainer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.MainSplitContainer.Panel1.SuspendLayout()
        Me.MainSplitContainer.SuspendLayout()
        Me.StatusStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'MainSplitContainer
        '
        Me.MainSplitContainer.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.MainSplitContainer.Location = New System.Drawing.Point(2, 35)
        Me.MainSplitContainer.Name = "MainSplitContainer"
        '
        'MainSplitContainer.Panel1
        '
        Me.MainSplitContainer.Panel1.Controls.Add(Me.ResourceExplorer)
        '
        'MainSplitContainer.Panel2
        '
        Me.MainSplitContainer.Panel2.BackColor = System.Drawing.Color.White
        Me.MainSplitContainer.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Zoom
        Me.MainSplitContainer.Size = New System.Drawing.Size(1012, 531)
        Me.MainSplitContainer.SplitterDistance = 243
        Me.MainSplitContainer.TabIndex = 0
        '
        'ResourceExplorer
        '
        Me.ResourceExplorer.AllowDrop = True
        Me.ResourceExplorer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ResourceExplorer.HideSelection = False
        Me.ResourceExplorer.Location = New System.Drawing.Point(0, 0)
        Me.ResourceExplorer.Name = "ResourceExplorer"
        TreeNode1.Name = "Computers"
        TreeNode1.Text = "Computers"
        TreeNode2.Name = "Collections"
        TreeNode2.Text = "Collections"
        TreeNode3.Name = "Task"
        TreeNode3.Text = "Task"
        TreeNode4.Name = "Query"
        TreeNode4.Text = "Query"
        TreeNode5.Name = "Results"
        TreeNode5.Text = "Results"
        TreeNode6.Name = "RootNode"
        TreeNode6.Text = "Resource Explorer"
        TreeNode7.Name = "CustomActions"
        TreeNode7.Text = "Custom Actions"
        TreeNode8.Name = "Settings"
        TreeNode8.Text = "Settings"
        Me.ResourceExplorer.Nodes.AddRange(New System.Windows.Forms.TreeNode() {TreeNode6, TreeNode8})
        Me.ResourceExplorer.Size = New System.Drawing.Size(243, 531)
        Me.ResourceExplorer.TabIndex = 0
        '
        'UserInputComboBox
        '
        Me.UserInputComboBox.AllowDrop = True
        Me.UserInputComboBox.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.UserInputComboBox.FormattingEnabled = True
        Me.UserInputComboBox.Location = New System.Drawing.Point(107, 8)
        Me.UserInputComboBox.Name = "UserInputComboBox"
        Me.UserInputComboBox.Size = New System.Drawing.Size(652, 21)
        Me.UserInputComboBox.TabIndex = 0
        '
        'LabelResource
        '
        Me.LabelResource.AutoSize = True
        Me.LabelResource.Location = New System.Drawing.Point(12, 11)
        Me.LabelResource.Name = "LabelResource"
        Me.LabelResource.Size = New System.Drawing.Size(89, 13)
        Me.LabelResource.TabIndex = 1
        Me.LabelResource.Text = "Select Resource:"
        '
        'SubmitButton
        '
        Me.SubmitButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.SubmitButton.Location = New System.Drawing.Point(765, 6)
        Me.SubmitButton.Name = "SubmitButton"
        Me.SubmitButton.Size = New System.Drawing.Size(75, 23)
        Me.SubmitButton.TabIndex = 2
        Me.SubmitButton.Text = "Submit"
        Me.SubmitButton.UseVisualStyleBackColor = True
        '
        'ClearButton
        '
        Me.ClearButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.ClearButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.ClearButton.Location = New System.Drawing.Point(846, 6)
        Me.ClearButton.Name = "ClearButton"
        Me.ClearButton.Size = New System.Drawing.Size(75, 23)
        Me.ClearButton.TabIndex = 3
        Me.ClearButton.Text = "Clear"
        Me.ClearButton.UseVisualStyleBackColor = True
        '
        'NewButton
        '
        Me.NewButton.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.NewButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.NewButton.Location = New System.Drawing.Point(927, 6)
        Me.NewButton.Name = "NewButton"
        Me.NewButton.Size = New System.Drawing.Size(75, 23)
        Me.NewButton.TabIndex = 4
        Me.NewButton.Text = "New"
        Me.NewButton.UseVisualStyleBackColor = True
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.UpTimeStatusLabel, Me.MemoryUsageStatusLabel, Me.CpuUsageStatusLabel, Me.ConnectionsStatusLabel, Me.DateTimeStatusLabel})
        Me.StatusStrip1.LayoutStyle = System.Windows.Forms.ToolStripLayoutStyle.HorizontalStackWithOverflow
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 569)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(1014, 22)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'UpTimeStatusLabel
        '
        Me.UpTimeStatusLabel.BorderStyle = System.Windows.Forms.Border3DStyle.Etched
        Me.UpTimeStatusLabel.Name = "UpTimeStatusLabel"
        Me.UpTimeStatusLabel.Size = New System.Drawing.Size(109, 17)
        Me.UpTimeStatusLabel.Text = "UpTimeStatusLabel"
        '
        'MemoryUsageStatusLabel
        '
        Me.MemoryUsageStatusLabel.Name = "MemoryUsageStatusLabel"
        Me.MemoryUsageStatusLabel.Size = New System.Drawing.Size(144, 17)
        Me.MemoryUsageStatusLabel.Text = "MemoryUsageStatusLabel"
        '
        'CpuUsageStatusLabel
        '
        Me.CpuUsageStatusLabel.Name = "CpuUsageStatusLabel"
        Me.CpuUsageStatusLabel.Size = New System.Drawing.Size(121, 17)
        Me.CpuUsageStatusLabel.Text = "CpuUsageStatusLabel"
        '
        'ConnectionsStatusLabel
        '
        Me.ConnectionsStatusLabel.Name = "ConnectionsStatusLabel"
        Me.ConnectionsStatusLabel.Size = New System.Drawing.Size(134, 17)
        Me.ConnectionsStatusLabel.Text = "ConnectionsStatusLabel"
        '
        'CurrentTimeStatusLabel
        '
        Me.DateTimeStatusLabel.Name = "CurrentTimeStatusLabel"
        Me.DateTimeStatusLabel.Size = New System.Drawing.Size(134, 17)
        Me.DateTimeStatusLabel.Text = "CurrentTimeStatusLabel"
        '
        'MainForm
        '
        Me.AcceptButton = Me.SubmitButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.ClearButton
        Me.ClientSize = New System.Drawing.Size(1014, 591)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.NewButton)
        Me.Controls.Add(Me.ClearButton)
        Me.Controls.Add(Me.SubmitButton)
        Me.Controls.Add(Me.LabelResource)
        Me.Controls.Add(Me.UserInputComboBox)
        Me.Controls.Add(Me.MainSplitContainer)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "MainForm"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "System Tool 3"
        Me.MainSplitContainer.Panel1.ResumeLayout(False)
        CType(Me.MainSplitContainer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.MainSplitContainer.ResumeLayout(False)
        Me.StatusStrip1.ResumeLayout(False)
        Me.StatusStrip1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents MainSplitContainer As System.Windows.Forms.SplitContainer
    Friend WithEvents UserInputComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents LabelResource As System.Windows.Forms.Label
    Friend WithEvents SubmitButton As System.Windows.Forms.Button
    Public WithEvents ResourceExplorer As System.Windows.Forms.TreeView
    Friend WithEvents ClearButton As System.Windows.Forms.Button
    Friend WithEvents NewButton As System.Windows.Forms.Button
    Friend WithEvents StatusStrip1 As StatusStrip
    Friend WithEvents UpTimeStatusLabel As ToolStripStatusLabel
    Friend WithEvents CpuUsageStatusLabel As ToolStripStatusLabel
    Friend WithEvents MemoryUsageStatusLabel As ToolStripStatusLabel
    Friend WithEvents ConnectionsStatusLabel As ToolStripStatusLabel
    Friend WithEvents DateTimeStatusLabel As ToolStripStatusLabel
End Class
