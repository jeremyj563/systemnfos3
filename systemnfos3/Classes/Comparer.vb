Public Class Comparer
    Implements IComparer

    Private Property ControlType As Type
    Private Property DateField As Boolean = False
    Private Property ListViewColumnIndex As Integer = 0
    Private Property SortType As SortTypes = SortTypes.Ascending
    Private Property SortAction As Boolean

    Public Sub New(listViewColumnIndex As Integer, dateField As Boolean, sortType As SortTypes)
        Me.ControlType = GetType(ListView)
        Me.ListViewColumnIndex = listViewColumnIndex
        Me.DateField = dateField
        Me.SortType = sortType
    End Sub

    Public Sub New(treeView As TreeView, sortAction As Boolean)
        Me.ControlType = GetType(TreeView)
        Me.SortAction = sortAction
    End Sub

    Public Function Compare(x As Object, y As Object) As Integer Implements IComparer.Compare
        Dim returnValue As Integer = -1

        Select Case Me.ControlType

            Case GetType(ListView)
                Select Case Me.DateField

                    Case True
                        Select Case Me.SortType

                            Case SortTypes.Ascending
                                returnValue = Date.Compare(CType(x, ListViewItem).SubItems(ListViewColumnIndex).Text, CType(y, ListViewItem).SubItems(ListViewColumnIndex).Text)
                                Return returnValue

                            Case SortTypes.Descending
                                returnValue = Date.Compare(CType(y, ListViewItem).SubItems(ListViewColumnIndex).Text, CType(x, ListViewItem).SubItems(ListViewColumnIndex).Text)
                                Return returnValue

                        End Select

                    Case False
                        Select Case Me.SortType

                            Case SortTypes.Ascending
                                returnValue = String.Compare(CType(x, ListViewItem).SubItems(ListViewColumnIndex).Text, CType(y, ListViewItem).SubItems(ListViewColumnIndex).Text)
                                Return returnValue

                            Case SortTypes.Descending
                                returnValue = String.Compare(CType(y, ListViewItem).SubItems(ListViewColumnIndex).Text, CType(x, ListViewItem).SubItems(ListViewColumnIndex).Text)
                                Return returnValue

                        End Select
                End Select

            Case GetType(TreeView)
                Select Case SortAction

                    Case True
                        Dim xNode As TreeNode = CType(x, TreeNode)
                        Dim yNode As TreeNode = CType(y, TreeNode)

                        If xNode.Tag Is Nothing Then Return 0

                        If xNode.Name = NameOf(Nodes.Settings) Then Return 0
                        If xNode.Name = NameOf(Nodes.RootNode) Then Return 0
                        If xNode.Parent.Name <> NameOf(Nodes.Computers) Then Return 0

                        Dim xComputer As ComputerControl = CType(xNode.Tag, ComputerControl)
                        Dim yComputer As ComputerControl = CType(yNode.Tag, ComputerControl)

                        If xComputer Is Nothing OrElse yComputer Is Nothing Then Return 0

                        If xComputer.ConnectionStatus = Nothing Then Return 0

                        If (xComputer.UserStatus + (xComputer.ConnectionStatus * 10)) <> (yComputer.UserStatus + (yComputer.ConnectionStatus * 10)) Then
                            Return xComputer.UserStatus + (xComputer.ConnectionStatus * 10) - (yComputer.UserStatus + (yComputer.ConnectionStatus * 10))
                        Else
                            Return String.Compare(xNode.Text, yNode.Text)
                        End If

                    Case False
                        Return 0

                End Select
        End Select

        Return returnValue
    End Function

    Public Enum SortTypes
        Ascending = 0
        Descending = 1
    End Enum

End Class