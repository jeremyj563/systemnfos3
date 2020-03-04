<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QueryControl
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
        Me.LDAPRadioButton = New System.Windows.Forms.RadioButton()
        Me.WMIRadioButton = New System.Windows.Forms.RadioButton()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.QueryNameTextBox = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.AddButton = New System.Windows.Forms.Button()
        Me.AttributeComboBox = New System.Windows.Forms.ComboBox()
        Me.AttributeClassComboBox = New System.Windows.Forms.ComboBox()
        Me.QueryListView = New System.Windows.Forms.ListView()
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader2 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.CollectionsTreeView = New System.Windows.Forms.TreeView()
        Me.StartButton = New System.Windows.Forms.Button()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.LDAPRadioButton)
        Me.GroupBox1.Controls.Add(Me.WMIRadioButton)
        Me.GroupBox1.Controls.Add(Me.Label5)
        Me.GroupBox1.Controls.Add(Me.QueryNameTextBox)
        Me.GroupBox1.Controls.Add(Me.Label4)
        Me.GroupBox1.Controls.Add(Me.Label1)
        Me.GroupBox1.Controls.Add(Me.AddButton)
        Me.GroupBox1.Controls.Add(Me.AttributeComboBox)
        Me.GroupBox1.Controls.Add(Me.AttributeClassComboBox)
        Me.GroupBox1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(207, 3)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(286, 203)
        Me.GroupBox1.TabIndex = 26
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Select Query Parameters"
        '
        'RadioButtonLdap
        '
        Me.LDAPRadioButton.AutoSize = True
        Me.LDAPRadioButton.Location = New System.Drawing.Point(102, 60)
        Me.LDAPRadioButton.Name = "RadioButtonLdap"
        Me.LDAPRadioButton.Size = New System.Drawing.Size(53, 17)
        Me.LDAPRadioButton.TabIndex = 21
        Me.LDAPRadioButton.TabStop = True
        Me.LDAPRadioButton.Text = "LDAP"
        Me.LDAPRadioButton.UseVisualStyleBackColor = True
        '
        'RadioButtonWMI
        '
        Me.WMIRadioButton.AutoSize = True
        Me.WMIRadioButton.Location = New System.Drawing.Point(6, 60)
        Me.WMIRadioButton.Name = "RadioButtonWMI"
        Me.WMIRadioButton.Size = New System.Drawing.Size(48, 17)
        Me.WMIRadioButton.TabIndex = 20
        Me.WMIRadioButton.TabStop = True
        Me.WMIRadioButton.Text = "WMI"
        Me.WMIRadioButton.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(6, 22)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(69, 13)
        Me.Label5.TabIndex = 19
        Me.Label5.Text = "Query Name:"
        '
        'TextBoxQueryName
        '
        Me.QueryNameTextBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QueryNameTextBox.Location = New System.Drawing.Point(81, 19)
        Me.QueryNameTextBox.Name = "TextBoxQueryName"
        Me.QueryNameTextBox.Size = New System.Drawing.Size(197, 20)
        Me.QueryNameTextBox.TabIndex = 0
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(3, 159)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(91, 13)
        Me.Label4.TabIndex = 15
        Me.Label4.Text = "Attribute Property:"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(6, 100)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(77, 13)
        Me.Label1.TabIndex = 12
        Me.Label1.Text = "Attribute Class:"
        '
        'ButtonAdd
        '
        Me.AddButton.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AddButton.Location = New System.Drawing.Point(236, 173)
        Me.AddButton.Name = "ButtonAdd"
        Me.AddButton.Size = New System.Drawing.Size(42, 23)
        Me.AddButton.TabIndex = 3
        Me.AddButton.Text = "Add"
        Me.AddButton.UseVisualStyleBackColor = True
        '
        'ComboBoxAttribute
        '
        Me.AttributeComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.AttributeComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AttributeComboBox.FormattingEnabled = True
        Me.AttributeComboBox.Location = New System.Drawing.Point(6, 175)
        Me.AttributeComboBox.Name = "ComboBoxAttribute"
        Me.AttributeComboBox.Size = New System.Drawing.Size(224, 21)
        Me.AttributeComboBox.Sorted = True
        Me.AttributeComboBox.TabIndex = 2
        '
        'ComboBoxAttClass
        '
        Me.AttributeClassComboBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.AttributeClassComboBox.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.AttributeClassComboBox.FormattingEnabled = True
        Me.AttributeClassComboBox.Location = New System.Drawing.Point(6, 116)
        Me.AttributeClassComboBox.Name = "ComboBoxAttClass"
        Me.AttributeClassComboBox.Size = New System.Drawing.Size(224, 21)
        Me.AttributeClassComboBox.Sorted = True
        Me.AttributeClassComboBox.TabIndex = 1
        '
        'ListViewQuery
        '
        Me.QueryListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.ColumnHeader1, Me.ColumnHeader2})
        Me.QueryListView.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.QueryListView.FullRowSelect = True
        Me.QueryListView.HeaderStyle = System.Windows.Forms.ColumnHeaderStyle.Nonclickable
        Me.QueryListView.Location = New System.Drawing.Point(207, 212)
        Me.QueryListView.Name = "ListViewQuery"
        Me.QueryListView.Size = New System.Drawing.Size(286, 280)
        Me.QueryListView.TabIndex = 20
        Me.QueryListView.UseCompatibleStateImageBehavior = False
        Me.QueryListView.View = System.Windows.Forms.View.Details
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "Properties to query for:"
        Me.ColumnHeader1.Width = 280
        '
        'ColumnHeader2
        '
        Me.ColumnHeader2.Text = "Query type:"
        Me.ColumnHeader2.Width = 67
        '
        'TreeViewCollections
        '
        Me.CollectionsTreeView.Location = New System.Drawing.Point(3, 3)
        Me.CollectionsTreeView.Name = "TreeViewCollections"
        Me.CollectionsTreeView.Size = New System.Drawing.Size(198, 489)
        Me.CollectionsTreeView.TabIndex = 29
        '
        'ButtonStart
        '
        Me.StartButton.Location = New System.Drawing.Point(207, 498)
        Me.StartButton.Name = "ButtonStart"
        Me.StartButton.Size = New System.Drawing.Size(71, 23)
        Me.StartButton.TabIndex = 30
        Me.StartButton.Text = "Start Query"
        Me.StartButton.UseVisualStyleBackColor = True
        '
        'ControlQuery
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.StartButton)
        Me.Controls.Add(Me.CollectionsTreeView)
        Me.Controls.Add(Me.QueryListView)
        Me.Controls.Add(Me.GroupBox1)
        Me.Name = "ControlQuery"
        Me.Size = New System.Drawing.Size(723, 524)
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents GroupBox1 As System.Windows.Forms.GroupBox
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents QueryNameTextBox As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents AddButton As System.Windows.Forms.Button
    Friend WithEvents AttributeComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents AttributeClassComboBox As System.Windows.Forms.ComboBox
    Friend WithEvents QueryListView As System.Windows.Forms.ListView
    Friend WithEvents LDAPRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents WMIRadioButton As System.Windows.Forms.RadioButton
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents CollectionsTreeView As System.Windows.Forms.TreeView
    Friend WithEvents StartButton As System.Windows.Forms.Button
    Friend WithEvents ColumnHeader2 As System.Windows.Forms.ColumnHeader

End Class
