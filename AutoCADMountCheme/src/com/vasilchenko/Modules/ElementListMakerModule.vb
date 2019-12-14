Imports System.Threading
Imports AutoCADMountCheme.com.vasilchenko.classes
Imports AutoCADMountCheme.com.vasilchenko.DatabaseConnection


Namespace com.vasilchenko.Modules
    Module ElementListMakerModule
        Sub BuildMountSchema()
            Dim acElementList As New List(Of  ElementClass)
            
            Dim acElementTagList As List(Of  String) = DatabaseDataAccessObject.GetAllElementTag()
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
