Imports AutoCADMountCheme.com.vasilchenko.classes
Imports AutoCADMountCheme.com.vasilchenko.DatabaseConnection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput

Namespace com.vasilchenko.Modules
    Public Module RedrawElementModule
        Public Sub RedrawElement()
            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor


            Const i As Short = 1
            While i <> 0
                Dim acPromptPntOpt As New PromptSelectionOptions()
                acPromptPntOpt.MessageForAdding = vbLf & "Выбери элемент для перерисовки"
                acPromptPntOpt.SingleOnly = true
                Dim acPromptSelRes As PromptSelectionResult = acEditor.GetSelection(acPromptPntOpt)

                If acPromptSelRes.Status <> PromptStatus.OK Then
                    acEditor.WriteMessage("Не выбран элемент" & vbCrLf)
                    Exit Sub
                Else
                    Using acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction
                        Dim acBlockRef As BlockReference = acTransaction.GetObject(acPromptSelRes.Value(0).ObjectId,
                                                                                   OpenMode.ForWrite)
                        Dim acCurElement as New ElementClass
                        For Each id As ObjectId In acBlockRef.AttributeCollection
                            Dim acAttrReference As AttributeReference = acTransaction.GetObject(id, OpenMode.ForRead)
                            Select Case acAttrReference.Tag
                                Case "P_TAG1"
                                    acCurElement.TagName = acAttrReference.TextString
                            End Select
                        Next

                        if IsNothing(acCurElement.TagName) Then
                            acEditor.WriteMessage("Тэг не найден" & vbCrLf)
                            Exit Sub
                        End If

                        Dim acBasePoint = acBlockRef.Position
                        acBlockRef.Erase()
                        acEditor.Regen()

                        acCurElement = DatabaseDataAccessObject.GetSingleElementData(acCurElement.TagName)
                        DatabaseDataAccessObject.FillElementData(acCurElement)
                        DatabaseDataAccessObject.FillBlockPath(acCurElement)
                        DatabaseDataAccessObject.FillElementConnections(acCurElement)

                        DrawingBlocksModule.DrawElement(acCurElement, acBasePoint)

                        acTransaction.Commit()
                    End Using
                End If
            End While
        End Sub
    End Module
End Namespace