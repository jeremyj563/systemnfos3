<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class CollectionNameForm
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
        Me.TableLayoutPanel1 = New System.Windows.Forms.TableLayoutPanel()
        Me.OKButton = New System.Windows.Forms.Button()
        Me.CancelButton = New System.Windows.Forms.Button()
        Me.ColumnNameTextBox = New System.Windows.Forms.TextBox()
        Me.TableLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'TableLayoutPanel1
        '
        Me.TableLayoutPanel1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.TableLayoutPanel1.ColumnCount = 2
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.ColumnStyles.Add(New System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Controls.Add(Me.OKButton, 0, 0)
        Me.TableLayoutPanel1.Controls.Add(Me.CancelButton, 1, 0)
        Me.TableLayoutPanel1.Location = New System.Drawing.Point(231, 6)
        Me.TableLayoutPanel1.Name = "TableLayoutPanel1"
        Me.TableLayoutPanel1.RowCount = 1
        Me.TableLayoutPanel1.RowStyles.Add(New System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Percent, 50.0!))
        Me.TableLayoutPanel1.Size = New System.Drawing.Size(146, 29)
        Me.TableLayoutPanel1.TabIndex = 0
        '
        'OK_Button
        '
        Me.OKButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.OKButton.Location = New System.Drawing.Point(3, 3)
        Me.OKButton.Name = "OK_Button"
        Me.OKButton.Size = New System.Drawing.Size(67, 23)
        Me.OKButton.TabIndex = 1
        Me.OKButton.Text = "OK"
        '
        'Cancel_Button
        '
        Me.CancelButton.Anchor = System.Windows.Forms.AnchorStyles.None
        Me.CancelButton.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.CancelButton.Location = New System.Drawing.Point(76, 3)
        Me.CancelButton.Name = "Cancel_Button"
        Me.CancelButton.Size = New System.Drawing.Size(67, 23)
        Me.CancelButton.TabIndex = 2
        Me.CancelButton.Text = "Cancel"
        '
        'txtColName
        '
        Me.ColumnNameTextBox.Location = New System.Drawing.Point(12, 12)
        Me.ColumnNameTextBox.Name = "txtColName"
        Me.ColumnNameTextBox.Size = New System.Drawing.Size(208, 20)
        Me.ColumnNameTextBox.TabIndex = 0
        '
        'dialColName
        '
        Me.AcceptButton = Me.OKButton
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.CancelButton
        Me.ClientSize = New System.Drawing.Size(389, 47)
        Me.Controls.Add(Me.ColumnNameTextBox)
        Me.Controls.Add(Me.TableLayoutPanel1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "DialColName"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "Enter Collection Name"
        Me.TableLayoutPanel1.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents TableLayoutPanel1 As System.Windows.Forms.TableLayoutPanel
    Friend WithEvents OKButton As System.Windows.Forms.Button
    Friend WithEvents CancelButton As System.Windows.Forms.Button
    Friend WithEvents ColumnNameTextBox As System.Windows.Forms.TextBox

End Class
