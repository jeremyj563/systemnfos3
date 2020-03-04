Imports System.IO
Imports Microsoft.Win32

Public Class CollectionController

    Public Shared Function GetAllCollections() As String()
        Return New RegistryController().GetKeyValues(My.Settings.RegistryPathCollections, RegistryHive.CurrentUser)
    End Function

    Public Shared Sub AddToCollection(collectionName As String, computerName As String)
        For Each collection As String In GetAllCollections()

            If collectionName.ToUpper() = collection.ToUpper() Then
                Dim exists As Boolean = False
                Dim collectionKey As String = Path.Combine(My.Settings.RegistryPathCollections, collection)

                Dim coll As String() = New RegistryController().GetKeyValues(collectionKey, RegistryHive.CurrentUser)
                For Each computerName In coll
                    If computerName.ToUpper().Trim() = computerName.ToUpper().Trim() Then
                        exists = True
                        Exit For
                    End If
                Next

                If Not exists Then
                    Dim registry As New RegistryController
                    registry.NewKey(collectionKey, RegistryHive.CurrentUser)
                End If

                Exit For
            End If
        Next
    End Sub

End Class