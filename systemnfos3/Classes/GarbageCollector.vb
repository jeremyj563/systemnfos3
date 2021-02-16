Imports System.Reflection
Imports System.ComponentModel

Public Class GarbageCollector

    Private Property GarbageCan As List(Of Object)
    Private garbageLock As New Object

    Public Sub New()
        Me.GarbageCan = Nothing
    End Sub

    Public WriteOnly Property AddToGarbage
        Set(Value)
            SyncLock garbageLock
                If Me.GarbageCan Is Nothing Then
                    Me.GarbageCan = New List(Of Object)()
                End If

                Me.GarbageCan.Add(Value)
            End SyncLock
        End Set
    End Property

    Public Sub DisposeAsync()
        Dim worker As New BackgroundWorker()
        AddHandler worker.DoWork, AddressOf CollectGarbage
        AddHandler worker.RunWorkerCompleted, Sub() DisposeGarbage(worker)

        worker.RunWorkerAsync()
    End Sub

    Private Sub CollectGarbage()
        SyncLock garbageLock
            If Me.GarbageCan IsNot Nothing Then
                For index As Integer = 0 To Me.GarbageCan.Count - 1
                    If Me.GarbageCan(index) IsNot Nothing AndAlso TypeOf Me.GarbageCan(index) Is ComputerPanel Then
                        Try
                            Dim panel As ComputerPanel = Me.GarbageCan(index)
                            With panel.InitWorker
                                .CancelAsync()
                                .Dispose()
                            End With

                            panel.UIThread(Sub() panel.Dispose())

                            Me.GarbageCan(index) = Nothing
                        Catch ex As InvalidOperationException
                            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                            Me.GarbageCan(index) = Nothing
                        Catch ex As Exception
                            LogEvent(String.Format("EXCEPTION in {0}: {1}", MethodBase.GetCurrentMethod(), ex.Message))
                        End Try
                    End If
                Next
            End If
        End SyncLock
    End Sub

    Private Sub DisposeGarbage(worker As BackgroundWorker)
        SyncLock garbageLock
            Me.GarbageCan = Nothing
        End SyncLock

        worker.Dispose()
    End Sub

End Class