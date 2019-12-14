Imports System.Threading
Imports AutoCADMountCheme.com.vasilchenko.classes
Imports AutoCADMountCheme.com.vasilchenko.DatabaseConnection


Namespace com.vasilchenko.Modules
    Module SingleElementListMakerModule
        Sub BuildMountElements(acElementTagList As List(Of  String))
            Dim acElementList As New List(Of  ElementClass)

            For Each element In acElementTagList
                DatabaseDataAccessObject.GetAllElementData(element, acElementList)
            Next
            
            For Each element In acElementList
                DatabaseDataAccessObject.FillElementData(element)
                DatabaseDataAccessObject.FillBlockPath(element)
                DatabaseDataAccessObject.FillElementConnections(element)
            Next

            DrawingBlocksModule.DrawBlocks(acElementList)
        End Sub
    End Module
End Namespace
