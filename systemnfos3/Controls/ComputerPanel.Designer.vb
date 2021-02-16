<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ComputerPanel
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
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
        Me.SplitContainer = New System.Windows.Forms.SplitContainer()
        Me.StatusRichTextBox = New System.Windows.Forms.RichTextBox()
        CType(Me.SplitContainer, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SplitContainer.Panel2.SuspendLayout()
        Me.SplitContainer.SuspendLayout()
        Me.SuspendLayout()
        '
        'SplitContainer
        '
        Me.SplitContainer.Dock = System.Windows.Forms.DockStyle.Fill
        Me.SplitContainer.Location = New System.Drawing.Point(0, 0)
        Me.SplitContainer.Name = "SplitContainer"
        Me.SplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal
        '
        'SplitContainer.Panel2
        '
        Me.SplitContainer.Panel2.Controls.Add(Me.StatusRichTextBox)
        Me.SplitContainer.Size = New System.Drawing.Size(648, 483)
        Me.SplitContainer.SplitterDistance = 418
        Me.SplitContainer.TabIndex = 0
        '
        'StatusRichTextBox
        '
        Me.StatusRichTextBox.Dock = System.Windows.Forms.DockStyle.Fill
        Me.StatusRichTextBox.Location = New System.Drawing.Point(0, 0)
        Me.StatusRichTextBox.Name = "StatusRichTextBox"
        Me.StatusRichTextBox.ReadOnly = True
        Me.StatusRichTextBox.Size = New System.Drawing.Size(648, 61)
        Me.StatusRichTextBox.TabIndex = 0
        Me.StatusRichTextBox.Text = ""
        '
        'ComputerControl
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.Controls.Add(Me.SplitContainer)
        Me.Name = "ComputerControl"
        Me.Size = New System.Drawing.Size(648, 483)
        Me.SplitContainer.Panel2.ResumeLayout(False)
        CType(Me.SplitContainer, System.ComponentModel.ISupportInitialize).EndInit()
        Me.SplitContainer.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents SplitContainer As System.Windows.Forms.SplitContainer
    Friend WithEvents StatusRichTextBox As System.Windows.Forms.RichTextBox

End Class
