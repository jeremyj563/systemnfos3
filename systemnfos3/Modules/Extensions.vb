Imports System.ComponentModel
Imports System.Reflection
Imports System.Runtime.CompilerServices
Imports System.Runtime.InteropServices

#Region " DataTableExtensions "

' Source: https://docs.microsoft.com/en-us/dotnet/framework/data/adonet/implement-copytodatatable-where-type-not-a-datarow
''' <summary>
''' Extension methods for <see cref="System.Data.DataTable"/>.
''' </summary>
Public Module DataTableExtensions

    <Extension()>
    Public Function CopyToDataTable(Of T)(ByVal source As IEnumerable(Of T)) As DataTable
        Return New ObjectShredder(Of T)().Shred(source, Nothing, Nothing)
    End Function

    <Extension()>
    Public Function CopyToDataTable(Of T)(ByVal source As IEnumerable(Of T), ByVal table As DataTable, ByVal options As LoadOption?) As DataTable
        Return New ObjectShredder(Of T)().Shred(source, table, options)
    End Function

    Private Class ObjectShredder(Of T)
        Private Property Fields As FieldInfo()
        Private Property OrdinalMap As Dictionary(Of String, Integer)
        Private Property Properties As PropertyInfo()
        Private Property Type As Type

        Public Sub New()
            Me.Type = GetType(T)
            Me.Fields = Me.Type.GetFields()
            Me.Properties = Me.Type.GetProperties()
            Me.OrdinalMap = New Dictionary(Of String, Integer)
        End Sub

        Public Function ShredObject(table As DataTable, instance As T) As Object()
            Dim fields As FieldInfo() = Me.Fields
            Dim properties As PropertyInfo() = Me.Properties

            If TypeOf instance IsNot T Then
                ' If the instance is derived from T, extend the table schema
                ' and get the properties and fields.
                Me.ExtendTable(table, instance.GetType())
                fields = instance.GetType().GetFields()
                properties = instance.GetType().GetProperties()
            End If

            ' Add the property and field values of the instance to an array.
            Dim values = New Object(table.Columns.Count - 1) {}
            For Each field As FieldInfo In fields
                values(Me.OrdinalMap.Item(field.Name)) = field.GetValue(instance)
            Next
            For Each [property] As PropertyInfo In properties
                values(Me.OrdinalMap.Item([property].Name)) = [property].GetValue(instance, Nothing)
            Next

            ' Return the property and field values of the instance.
            Return values
        End Function

        ''' <summary>Loads a DataTable from a sequence of objects.</summary>
        ''' <param name="source">The sequence of objects to load into the DataTable.</param>
        ''' <param name="table">The input table. The schema of the table must match that the type T. If the table is null, a new table is created with a schema created from the public properties and fields of the type T.</param>
        ''' <param name="options">Specifies how values from the source sequence will be applied to existing rows in the table.</param>
        ''' <returns>A DataTable created from the source sequence.</returns>
        Public Function Shred(source As IEnumerable(Of T), table As DataTable, options As LoadOption?) As DataTable

            ' Load the table from the scalar sequence if T is a primitive type.
            If GetType(T).IsPrimitive Then
                Return Me.ShredPrimitive(source, table, options)
            End If

            ' Create a new table if the input table is null.
            If table Is Nothing Then
                table = New DataTable(GetType(T).Name)
            End If

            ' Initialize the ordinal map and extend the table schema based on type T.
            table = Me.ExtendTable(table, GetType(T))

            ' Enumerate the source sequence and load the object values into rows.
            table.BeginLoadData()
            Using e As IEnumerator(Of T) = source.GetEnumerator()
                Do While e.MoveNext()
                    If options.HasValue Then
                        table.LoadDataRow(Me.ShredObject(table, e.Current), options.Value)
                    Else
                        table.LoadDataRow(Me.ShredObject(table, e.Current), True)
                    End If
                Loop
            End Using
            table.EndLoadData()

            ' Return the table.
            Return table
        End Function

        Public Function ShredPrimitive(source As IEnumerable(Of T), table As DataTable, options As LoadOption?) As DataTable
            ' Create a new table if the input table is null.
            If table Is Nothing Then
                table = New DataTable(GetType(T).Name)
            End If
            If Not table.Columns.Contains("Value") Then
                table.Columns.Add("Value", GetType(T))
            End If

            ' Enumerate the source sequence and load the scalar values into rows.
            table.BeginLoadData()

            Using e As IEnumerator(Of T) = source.GetEnumerator()
                Dim values = New Object(table.Columns.Count - 1) {}
                Do While e.MoveNext()
                    values(table.Columns("Value").Ordinal) = e.Current
                    If options.HasValue Then
                        table.LoadDataRow(values, options.Value)
                    Else
                        table.LoadDataRow(values, True)
                    End If
                Loop
            End Using

            table.EndLoadData()

            ' Return the table.
            Return table
        End Function

        Public Function ExtendTable(table As DataTable, type As Type) As DataTable
            ' Extend the table schema if the input table was null or if the value 
            ' in the sequence is derived from type T.
            Dim f As FieldInfo
            Dim p As PropertyInfo

            For Each f In type.GetFields
                If Not Me.OrdinalMap.ContainsKey(f.Name) Then
                    Dim dc As DataColumn

                    ' Add the field as a column in the table if it doesn't exist already.
                    dc = IIf(table.Columns.Contains(f.Name), table.Columns.Item(f.Name), table.Columns.Add(f.Name, f.FieldType))

                    ' Add the field to the ordinal map.
                    Me.OrdinalMap.Add(f.Name, dc.Ordinal)
                End If

            Next

            For Each p In type.GetProperties
                If Not Me.OrdinalMap.ContainsKey(p.Name) Then
                    ' Add the property as a column in the table if it doesn't exist already.
                    Dim dc As DataColumn
                    dc = IIf(table.Columns.Contains(p.Name), table.Columns.Item(p.Name), table.Columns.Add(p.Name, p.PropertyType))

                    ' Add the property to the ordinal map.
                    Me.OrdinalMap.Add(p.Name, dc.Ordinal)
                End If
            Next

            Return table
        End Function

    End Class

End Module

#End Region

#Region " BindingListExtensions "

' Source: https://stackoverflow.com/questions/43331145/how-can-i-improve-performance-of-an-addrange-method-on-a-custom-bindinglist

''' <summary>
''' Extension methods for <see cref="System.ComponentModel.BindingList(Of T)"/>.
''' </summary>
Public Module BindingListExtensions
    ''' <summary>
    ''' Adds the elements of the specified collection to the end of the <see cref="System.ComponentModel.BindingList(Of T)"/>,
    ''' while only firing the <see cref="System.ComponentModel.BindingList(Of T).ListChanged"/>-event once.
    ''' </summary>
    ''' <typeparam name="T">
    ''' The type T of the values of the <see cref="System.ComponentModel.BindingList(Of T)"/>.
    ''' </typeparam>
    ''' <param name="list">
    ''' The <see cref="System.ComponentModel.BindingList(Of T)"/> to which the values shall be added.
    ''' </param>
    ''' <param name="collection">
    ''' The collection whose elements should be added to the end of the <see cref="System.ComponentModel.BindingList(Of T)"/>.
    ''' The collection itself cannot be null, but it can contain elements that are null,
    ''' if type T is a reference type.
    ''' </param>
    ''' <exception cref="ArgumentNullException">values is null.</exception>
    <Extension()>
    Public Sub AddRange(Of T)(list As BindingList(Of T), collection As IEnumerable(Of T))
        ' The given collection may not be null
        If collection Is Nothing Then
            Throw New ArgumentNullException(NameOf(collection))
        End If

        ' Remember the current setting for RaiseListChangedEvents.
        ' If it was already deactivated we shouldn't activate it after adding.
        Dim originalRaiseEventsValue As Boolean = list.RaiseListChangedEvents

        ' Try adding all of the elements to the binding list.
        Try
            list.RaiseListChangedEvents = False

            For Each value As T In collection
                list.Add(value)
            Next
        Finally
            ' Restore the original setting for RaiseListChangedEvents (even if there was an exception)
            ' and fire the ListChanged-event once (if RaiseListChangedEvents is activated).
            list.RaiseListChangedEvents = originalRaiseEventsValue
            If list.RaiseListChangedEvents Then
                list.ResetBindings()
            End If
        End Try
    End Sub

End Module

#End Region

#Region " FormExtensions "
' Source: http://www.vbforums.com/showthread.php?367786-Flashing-Taskbar-(Like-MSN-Messanger)-Resolved
' Source: https://stackoverflow.com/questions/11309827/window-application-flash-like-orange-on-taskbar-when-minimize

Module FormExtensions
    Private Declare Function FlashWindowEx Lib "user32.dll" (ByRef pwfi As FLASHWINFO) As Boolean
    Private Const FLASHW_CAPTION As Int32 = &H1
    Private Const FLASHW_TRAY As Int32 = &H2
    Private Const FLASHW_ALL As Int32 = (FLASHW_CAPTION Or FLASHW_TRAY)

    <StructLayout(LayoutKind.Sequential)>
    Private Structure FLASHWINFO
        Public cbSize As UInt32
        Public hwnd As IntPtr
        Public dwFlags As UInt32
        Public uCount As UInt32
        Public dwTimeout As UInt32
    End Structure

    <Extension()>
    Public Function FlashNotification(form As Form, count As Integer) As Boolean
        Dim fInfo As New FLASHWINFO()

        With fInfo
            .cbSize = Convert.ToUInt32(Marshal.SizeOf(fInfo)) ' Size of the structure in bytes
            .hwnd = form.Handle     ' Handle of the window to flash
            .dwFlags = FLASHW_ALL   ' Flash *both* the caption bar and the tray
            .uCount = count         ' The number of flashes
            .dwTimeout = 1000       ' Speed of flashes in milliseconds (optional)
        End With

        Return FlashWindowEx(fInfo)
    End Function
End Module

#End Region

#Region " ControlExtensions "

Module ControlExtensions
    ' Source: https://www.codeproject.com/Articles/37642/Avoiding-InvokeRequired
    ' More info: http://www.interact-sw.co.uk/iangblog/2004/04/20/whatlocks

    <Extension()>
    Public Sub UIThread(control As Control, action As Action)
        If control.InvokeRequired Then
            control.SafeInvoke(action)
        Else
            action.SafeInvoke(control)
        End If
    End Sub

    <Extension()>
    Public Sub SafeInvoke(control As Control, action As Action)
        If Not control.IsDisposed AndAlso Not control.Disposing Then
            control.Invoke(action)
        End If
    End Sub

    <Extension()>
    Public Sub InvokeSetText(control As Control, text As String)
        control.UIThread(Sub() control.Text = text)
    End Sub

    <Extension()>
    Public Sub InvokeClearControls(control As Control)
        control.UIThread(Sub() control.Controls.Clear())
    End Sub

    <Extension()>
    Public Sub InvokeAddControl(control As Control, item As Control)
        control.UIThread(Sub() control.Controls.Add(item))
    End Sub

    <Extension()>
    Public Sub InvokeRemoveControl(control As Control, item As Control)
        control.UIThread(Sub() control.Controls.Remove(item))
    End Sub

    <Extension()>
    Public Sub InvokeCenterControl(control As Control)
        control.UIThread(Sub()
                             Dim center As New Point((control.Width / 2) - (control.Controls(0).Width / 2),
                                                     (control.Height / 2) - (control.Controls(0).Height / 2))

                             control.Controls(0).Location = center
                         End Sub)
    End Sub

End Module

#End Region

#Region "ActionExtensions"
Module ActionExtensions
    <Extension()>
    Public Sub SafeInvoke(action As Action, control As Control)
        If Not control.IsDisposed AndAlso Not control.Disposing Then
            action.Invoke()
        End If
    End Sub
End Module
#End Region