<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class QueryResultControl
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
        Me.lblName = New System.Windows.Forms.Label()
        Me.ResultsListView = New System.Windows.Forms.ListView()
        Me.SystemName = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.SuspendLayout()
        '
        'lblName
        '
        Me.lblName.AutoSize = True
        Me.lblName.Location = New System.Drawing.Point(3, 5)
        Me.lblName.Name = "lblName"
        Me.lblName.Size = New System.Drawing.Size(39, 13)
        Me.lblName.TabIndex = 5
        Me.lblName.Text = "Label1"
        '
        'lvResults
        '
        Me.ResultsListView.Columns.AddRange(New System.Windows.Forms.ColumnHeader() {Me.SystemName})
        Me.ResultsListView.FullRowSelect = True
        Me.ResultsListView.GridLines = True
        Me.ResultsListView.Location = New System.Drawing.Point(3, 21)
        Me.ResultsListView.Name = "lvResults"
        Me.ResultsListView.Size = New System.Drawing.Size(717, 500)
        Me.ResultsListView.TabIndex = 4
        Me.ResultsListView.UseCompatibleStateImageBehavior = False
        Me.ResultsListView.View = System.Windows.Forms.View.Details
        '
        'SystemName
        '
        Me.SystemName.Text = "System Name"
        Me.SystemName.Width = 120
        '
        'msQueryResult
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.lblName)
        Me.Controls.Add(Me.ResultsListView)
        Me.Name = "msQueryResult"
        Me.Size = New System.Drawing.Size(723, 524)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents lblName As System.Windows.Forms.Label
    Friend WithEvents ResultsListView As System.Windows.Forms.ListView
    Friend WithEvents SystemName As System.Windows.Forms.ColumnHeader

End Class
