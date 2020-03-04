<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class TaskResultControl
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
        Me.LabelName = New System.Windows.Forms.Label()
        Me.ResultsListView = New System.Windows.Forms.ListView()
        Me.SystemNames = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.Status = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'lblName
        '
        Me.LabelName.AutoSize = True
        Me.LabelName.Location = New System.Drawing.Point(3, 7)
        Me.LabelName.Name = "lblName"
        Me.LabelName.Size = New System.Drawing.Size(39, 13)
        Me.LabelName.TabIndex = 3
        Me.LabelName.Text = "Label1"
        '
        'lvResults
        '
        Me.ResultsListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.SystemNames, Me.Status})
        Me.ResultsListView.FullRowSelect = True
        Me.ResultsListView.GridLines = True
        Me.ResultsListView.Location = New System.Drawing.Point(3, 23)
        Me.ResultsListView.Name = "lvResults"
        Me.ResultsListView.Size = New System.Drawing.Size(717, 498)
        Me.ResultsListView.TabIndex = 2
        Me.ResultsListView.UseCompatibleStateImageBehavior = False
        Me.ResultsListView.View = System.Windows.Forms.View.Details
        '
        'SystemNames
        '
        Me.SystemNames.Text = "System Name"
        Me.SystemNames.Width = 120
        '
        'Status
        '
        Me.Status.Text = "Status"
        Me.Status.Width = 590
        '
        'msTaskResult
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.LabelName)
        Me.Controls.Add(Me.ResultsListView)
        Me.Name = "msTaskResult"
        Me.Size = New System.Drawing.Size(723, 524)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents LabelName As System.Windows.Forms.Label
    Friend WithEvents ResultsListView As System.Windows.Forms.ListView
    Friend WithEvents SystemNames As System.Windows.Forms.ColumnHeader
    Friend WithEvents Status As System.Windows.Forms.ColumnHeader

End Class
