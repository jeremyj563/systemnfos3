<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class CollectionOptionsControl
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
        Me.ComputerStatusListView = New System.Windows.Forms.ListView()
        Me.ComputerName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.AddButton = New System.Windows.Forms.Button()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.StatusTwoLabel = New System.Windows.Forms.Label()
        Me.SystemTextBox = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.ComputerNamesListView = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.StatusLabel = New System.Windows.Forms.Label()
        Me.ImportButton = New System.Windows.Forms.Button()
        Me.ButtonSelectList = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.SuspendLayout()
        '
        'lvSysStatus
        '
        Me.ComputerStatusListView.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ComputerStatusListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ComputerName})
        Me.ComputerStatusListView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComputerStatusListView.FullRowSelect = True
        Me.ComputerStatusListView.Location = New System.Drawing.Point(3, 3)
        Me.ComputerStatusListView.Name = "lvSysStatus"
        Me.ComputerStatusListView.Size = New System.Drawing.Size(176, 518)
        Me.ComputerStatusListView.TabIndex = 0
        Me.ComputerStatusListView.UseCompatibleStateImageBehavior = False
        Me.ComputerStatusListView.View = System.Windows.Forms.View.Details
        '
        'SystemName
        '
        Me.ComputerName.Text = "Collection Systems:"
        Me.ComputerName.Width = 171
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.AddButton)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.StatusTwoLabel)
        Me.GroupBox1.Controls.Add(Me.SystemTextBox)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(185, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(271, 86)
        Me.GroupBox1.TabIndex = 6
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Add single system to collection"
        '
        'btnAdd
        '
        Me.AddButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddButton.Location = New System.Drawing.Point(216, 30)
        Me.AddButton.Name = "btnAdd"
        Me.AddButton.Size = New System.Drawing.Size(49, 23)
        Me.AddButton.TabIndex = 11
        Me.AddButton.Text = "Add"
        Me.AddButton.UseVisualStyleBackColor = True
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 16)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(119, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Input a computer name:"
        '
        'lblStatusTwo
        '
        Me.StatusTwoLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusTwoLabel.Location = New System.Drawing.Point(6, 56)
        Me.StatusTwoLabel.Name = "lblStatusTwo"
        Me.StatusTwoLabel.Size = New System.Drawing.Size(259, 25)
        Me.StatusTwoLabel.TabIndex = 12
        '
        'txtSystem
        '
        Me.SystemTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.SystemTextBox.Location = New System.Drawing.Point(9, 32)
        Me.SystemTextBox.Name = "txtSystem"
        Me.SystemTextBox.Size = New System.Drawing.Size(201, 20)
        Me.SystemTextBox.TabIndex = 1
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.Label2)
        Me.GroupBox2.Controls.Add(Me.ComputerNamesListView)
        Me.GroupBox2.Controls.Add(Me.StatusLabel)
        Me.GroupBox2.Controls.Add(Me.ImportButton)
        Me.GroupBox2.Controls.Add(Me.ButtonSelectList)
        Me.GroupBox2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(185, 95)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(271, 206)
        Me.GroupBox2.TabIndex = 12
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "Import multiple systems from a text file"
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(6, 16)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(254, 40)
        Me.Label2.TabIndex = 6
        Me.Label2.Text = "Select a text file with a list of system names by clicking on 'Select List'. Then" & _
    " click 'Import Checked' to import the systems to your collection"
        '
        'lvSystemNames
        '
        Me.ComputerNamesListView.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.ComputerNamesListView.CheckBoxes = True
        Me.ComputerNamesListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1})
        Me.ComputerNamesListView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ComputerNamesListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.None
        Me.ComputerNamesListView.Location = New System.Drawing.Point(9, 59)
        Me.ComputerNamesListView.Name = "lvSystemNames"
        Me.ComputerNamesListView.Size = New System.Drawing.Size(201, 112)
        Me.ComputerNamesListView.TabIndex = 5
        Me.ComputerNamesListView.UseCompatibleStateImageBehavior = False
        Me.ComputerNamesListView.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Systems:"
        Me.ColumnHeader1.Width = 165
        '
        'lblStatus
        '
        Me.StatusLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.StatusLabel.Location = New System.Drawing.Point(173, 177)
        Me.StatusLabel.Name = "lblStatus"
        Me.StatusLabel.Size = New System.Drawing.Size(91, 23)
        Me.StatusLabel.TabIndex = 8
        Me.StatusLabel.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'btnImport
        '
        Me.ImportButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ImportButton.Location = New System.Drawing.Point(77, 177)
        Me.ImportButton.Name = "btnImport"
        Me.ImportButton.Size = New System.Drawing.Size(90, 23)
        Me.ImportButton.TabIndex = 7
        Me.ImportButton.Text = "Import Checked"
        Me.ImportButton.UseVisualStyleBackColor = True
        '
        'btnSelectList
        '
        Me.ButtonSelectList.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.ButtonSelectList.Location = New System.Drawing.Point(6, 177)
        Me.ButtonSelectList.Name = "btnSelectList"
        Me.ButtonSelectList.Size = New System.Drawing.Size(65, 23)
        Me.ButtonSelectList.TabIndex = 2
        Me.ButtonSelectList.Text = "Select List"
        Me.ButtonSelectList.UseVisualStyleBackColor = True
        '
        'msControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.ComputerStatusListView)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "msControl"
        Me.Size = New System.Drawing.Size(726, 524)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ComputerStatusListView As System.Windows.Forms.ListView
    Friend WithEvents ComputerName As System.Windows.Forms.ColumnHeader
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents AddButton As System.Windows.Forms.Button
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents StatusTwoLabel As System.Windows.Forms.Label
    Friend WithEvents SystemTextBox As System.Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As System.Windows.Forms.GroupBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents ComputerNamesListView As System.Windows.Forms.ListView
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ButtonSelectList As System.Windows.Forms.Button
    Friend WithEvents ImportButton As System.Windows.Forms.Button
    Friend WithEvents StatusLabel As System.Windows.Forms.Label

End Class
