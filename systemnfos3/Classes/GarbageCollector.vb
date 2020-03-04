Imports System.Reflection
Imports System.ComponentModel

Public Class GarbageCollector

    Private Property GarbageCan As List(Of Object)
    Private Property GarbageLock As New Object

    Public Sub New()
        Me.GarbageCan = Nothing
    End Sub

    Public WriteOnly Property AddToGarbage
        Set(Value)
            SyncLock Me.GarbageLock
                If Me.GarbageCan Is Nothing Then
                    Me.GarbageCan = New List(Of Object)()
                End If

                Me.GarbageCan.Add(Value)
            End SyncLock
        End Set
    End Property

    Public Sub DisposeAsync()
        Dim garbageCollectionWorker As New BackgroundWorker()
        AddHandler garbageCollectionWorker.DoWork, AddressOf CollectGarbage
        AddHandler garbageCollectionWorker.RunWorkerCompleted, Sub() DisposeGarbage(garbageCollectionWorker)

        garbageCollectionWorker.RunWorkerAsync()
    End Sub

    Private Sub CollectGarbage()
        SyncLock Me.GarbageLock
            If Me.GarbageCan IsNot Nothing Then
                For index As Integer = 0 To Me.GarbageCan.Count - 1
                    If Me.GarbageCan(index) IsNot Nothing AndAlso TypeOf Me.GarbageCan(index) Is ComputerControl Then
                        Try
                            Dim trashItem As ComputerControl = Me.GarbageCan(index)
                            With trashItem.LoaderBackgroundThread
                                .CancelAsync()
                                .Dispose()
                            End With

                            trashItem.UIThread(Sub() trashItem.Dispose())

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

    Private Sub DisposeGarbage(garbageCollectionWorker As BackgroundWorker)
        SyncLock GarbageLock
            Me.GarbageCan = Nothing
        End SyncLock

        garbageCollectionWorker.Dispose()
    End Sub

End Class