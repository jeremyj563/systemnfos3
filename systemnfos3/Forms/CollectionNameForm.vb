Public Class CollectionNameForm

    Private Property ColumnName As String = Nothing

    Private Sub OKButton_Click(sender As Object, e As EventArgs) Handles OKButton.Click
        If String.IsNullOrWhiteSpace(Me.ColumnNameTextBox.Text) Then
            MsgBox("You must enter a valid collection name!")
            Exit Sub
        Else
            Me.ColumnName = Me.ColumnNameTextBox.Text
            Me.DialogResult = DialogResult.OK
            Me.Close()
        End If
    End Sub

    Private Sub CancelButton_Click(sender As Object, e As EventArgs) Handles CancelButton.Click
        Me.DialogResult = DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CollectionNameForm_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.ColumnNameTextBox.Focus()
    End Sub
End Class
