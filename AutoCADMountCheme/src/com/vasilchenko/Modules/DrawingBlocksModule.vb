Imports AutoCADMountCheme.com.vasilchenko.classes
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.AutoCAD.DatabaseServices
Imports Autodesk.AutoCAD.EditorInput
Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.Modules
    Module DrawingBlocksModule
        Friend Sub DrawBlocks(acElementList As List(Of ElementClass))
            For Each element In acElementList
                DrawElement(element)
            Next
        End Sub

        Friend Sub DrawElement(element As ElementClass, Optional acBasePoint As Point3d = Nothing)

            Dim acDocument As Document = Application.DocumentManager.MdiActiveDocument
            Dim acDatabase As Database = acDocument.Database
            Dim acEditor As Editor = acDocument.Editor


            If Not IO.File.Exists(element.WdBlockPath) Then
                acEditor.WriteMessage(
                    "Для елемента " & element.CatalogName & " не сущесвует компоновочного образа" & vbCrLf)
                Exit Sub
            End If

            acEditor.WriteMessage("Рисуем " & element.TagName & vbCrLf)

            If acBasePoint = New Point3d(0,0,0) then
                Dim acPromptPntOpt As New PromptPointOptions("Выберите точку вставки")
                Dim acPromptPntResult As PromptPointResult = acEditor.GetPoint(acPromptPntOpt)

                If acPromptPntResult.Status <> PromptStatus.OK Then
                    acEditor.WriteMessage("Отмена вставки" & vbCrLf)
                    Exit Sub
                End If

                acBasePoint = acPromptPntResult.Value
            End If
            
            Dim acTransaction As Transaction = acDatabase.TransactionManager.StartTransaction()

            DrawLayerChecker.CheckLayers()

            Try
                Dim strBlkName As String = SymbolUtilityServices.GetBlockNameFromInsertPathName(element.WdBlockPath)
                Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
                Dim acInsObjectId As ObjectId

                If acBlockTable.Has(strBlkName) Then
                    Dim acCurrBlkTblRcd As BlockTableRecord = acTransaction.GetObject(acBlockTable.Item(strBlkName),
                                                                                      OpenMode.ForRead)
                    acInsObjectId = acCurrBlkTblRcd.Id
                Else
                    Dim acNewDbDwg As New Database(False, True)
                    acNewDbDwg.ReadDwgFile(element.WdBlockPath, FileOpenMode.OpenTryForReadShare, True, "")
                    acInsObjectId = acDatabase.Insert(strBlkName, acNewDbDwg, True)
                    acNewDbDwg.Dispose()
                End If

                Using acBlkRef As New BlockReference(acBasePoint, acInsObjectId)
                    acBlkRef.Layer = "PSYMS"
                    acBlkRef.ScaleFactors = New Scale3d(1)

                    Dim acBlockTableRecord As BlockTableRecord =
                            acTransaction.GetObject(acBlockTable.Item(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

                    acBlockTableRecord.AppendEntity(acBlkRef)
                    acTransaction.AddNewlyCreatedDBObject(acBlkRef, True)

                    Dim acBlockTableAttrRec As BlockTableRecord = acTransaction.GetObject(acInsObjectId,
                                                                                          OpenMode.ForRead)
                    Dim acAttrObjectId As ObjectId

                    For Each acAttrObjectId In acBlockTableAttrRec
                        Dim acAttrEntity As Entity = acTransaction.GetObject(acAttrObjectId, OpenMode.ForRead)
                        Dim acAttrDefinition = TryCast(acAttrEntity, AttributeDefinition)
                        If (acAttrDefinition IsNot Nothing) Then
                            Dim acAttrReference As New AttributeReference()
                            acAttrReference.SetAttributeFromBlock(acAttrDefinition, acBlkRef.BlockTransform)
                            Select Case acAttrReference.Tag
                                Case "P_TAG1"
                                    acAttrReference.TextString = element.TagName
                                    acAttrReference.TextStyleId =
                                        CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead),
                                              TextStyleTable)("WD_IEC")
                                    acAttrReference.Layer = "PTAG"
                                Case "MFG"
                                    acAttrReference.TextString = element.Manufacture
                                    acAttrReference.TextStyleId =
                                        CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead),
                                              TextStyleTable)("WD_IEC")
                                    acAttrReference.Layer = "PMFG"
                                Case "CAT"
                                    acAttrReference.TextString = element.CatalogName
                                    acAttrReference.TextStyleId =
                                        CType(acTransaction.GetObject(acDatabase.TextStyleTableId, OpenMode.ForRead),
                                              TextStyleTable)("WD_IEC")
                                    acAttrReference.Layer = "PCAT"
                                Case Else
                                    FillTermData(acAttrReference, element, acDatabase, acTransaction)

                            End Select
                            acBlkRef.AttributeCollection.AppendAttribute(acAttrReference)
                            acTransaction.AddNewlyCreatedDBObject(acAttrReference, True)
                        End If
                    Next
                End Using
                acTransaction.Commit()
            Catch ex As Exception
                MsgBox(
                    "ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr &
                    "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                acTransaction.Abort()
            Finally
                acTransaction.Dispose()
            End Try
        End Sub

        Private Sub FillTermData(acAttrReference As AttributeReference, element As ElementClass, acDatabase As Database,
                                 acTransaction As Transaction)

            For i = 0 To element.TermList.Count - 1
                Dim strI = ""
                If i < 9 Then strI = "0" & (i + 1) Else strI = i + 1

                If element.TermList.Values(i).ConnectionsList.Count > 2 Then
                    MsgBox(
                        "Для контакта " & element.TagName & ":" & element.TermList.Values(i).Term &
                        "слишком много соединений")
                End If

                Dim blnHasConnections = element.TermList.Values(i).ConnectionsList.Count <> 0

                Select Case acAttrReference.Tag
                    Case "TERM" & strI
                        acAttrReference.TextString = element.TermList.Values(i).Term
                        acAttrReference.Layer = "PTERM"
                    Case "WIRENO" & strI
                        If blnHasConnections Then
                            If element.TermList.Values(i).HasCableInConnection(0) Then
                                acAttrReference.TextString = element.TermList.Values(i).GetCableInConnection(0)
                            Else
                                acAttrReference.TextString = element.TermList.Values(i).GetConnections(0) &
                                                             element.TermList.Values(i).WireType
                            End If
                            acAttrReference.Layer = "PWIRE"
                        End If
                    Case "TERMDESC" & strI
                        If element.TermList.Values(i).ConnectionsList.Count > 1 Then
                            If element.TermList.Values(i).HasCableInConnection(0) Then
                                acAttrReference.TextString = element.TermList.Values(i).GetCableInConnection(1)
                            Else
                                acAttrReference.TextString = element.TermList.Values(i).GetConnections(1) &
                                                             element.TermList.Values(i).WireType
                            End If
                            acAttrReference.Layer = "PWIRE"
                        End If
                End Select

                If _
                    blnHasConnections AndAlso
                    (acAttrReference.Tag.StartsWith("X") And acAttrReference.Tag.EndsWith(strI)) Then
                    Select Case acAttrReference.Tag(1).ToString
                        Case "1"
                            element.TermList.Values(i).Direction = Enums.Direction.Right
                        Case "2"
                            element.TermList.Values(i).Direction = Enums.Direction.Top
                        Case "4"
                            element.TermList.Values(i).Direction = Enums.Direction.Left
                        Case "8"
                            element.TermList.Values(i).Direction = Enums.Direction.Bottom
                    End Select
                    element.TermList.Values(i).BasePoint = acAttrReference.Position
                    DrawLines(acDatabase, acTransaction, element.TermList.Values(i))
                End If
            Next
        End Sub

        Private Sub DrawLines(acDatabase As Database, acTransaction As Transaction, termClass As TermClass)
            Dim acBlockTable As BlockTable = acTransaction.GetObject(acDatabase.BlockTableId, OpenMode.ForRead)
            Dim acBlockTableRecord As BlockTableRecord =
                    acTransaction.GetObject(acBlockTable(BlockTableRecord.ModelSpace), OpenMode.ForWrite)

            Select Case termClass.Direction
                Case Enums.Direction.Bottom
                    DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, 0, - 30)
                    If termClass.ConnectionsList.Count > 1 Then
                        DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, - 7.5, - 1.5)
                        Dim acCurBasePoint = New Point3d(termClass.BasePoint.X - 7.5, termClass.BasePoint.Y - 1.5,
                                                         termClass.BasePoint.Z)
                        DrawSingleLine(acTransaction, acBlockTableRecord, acCurBasePoint, 0, - 30)
                    End If
                Case Enums.Direction.Left
                    DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, - 30, 0)
                    If termClass.ConnectionsList.Count > 1 Then
                        DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, - 1.5, 7.5)
                        Dim acCurBasePoint = New Point3d(termClass.BasePoint.X - 1.5, termClass.BasePoint.Y + 7.5,
                                                         termClass.BasePoint.Z)
                        DrawSingleLine(acTransaction, acBlockTableRecord, acCurBasePoint, - 30, 0)
                    End If
                Case Enums.Direction.Right
                    DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, 30, 0)
                    If termClass.ConnectionsList.Count > 1 Then
                        DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, 1.5, 7.5)
                        Dim acCurBasePoint = New Point3d(termClass.BasePoint.X + 1.5, termClass.BasePoint.Y + 7.5,
                                                         termClass.BasePoint.Z)
                        DrawSingleLine(acTransaction, acBlockTableRecord, acCurBasePoint, 30, 0)
                    End If
                Case Enums.Direction.Top
                    DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, 0, 30)
                    If termClass.ConnectionsList.Count > 1 Then
                        DrawSingleLine(acTransaction, acBlockTableRecord, termClass.BasePoint, - 7.5, 1.5)
                        Dim acCurBasePoint = New Point3d(termClass.BasePoint.X - 7.5, termClass.BasePoint.Y + 1.5,
                                                         termClass.BasePoint.Z)
                        DrawSingleLine(acTransaction, acBlockTableRecord, acCurBasePoint, 0, 30)
                    End If
            End Select
        End Sub

        Private Sub DrawSingleLine(acTransaction As Transaction, acBlockTableRecord As BlockTableRecord,
                                   basePoint As Point3d, dX As Double, dY As Double)
            Using acLine = New Line(BasePoint, New Point3d(BasePoint.X + dX, BasePoint.Y + dY, 0))
                acBlockTableRecord.AppendEntity(acLine)
                acTransaction.AddNewlyCreatedDBObject(acLine, True)
            End Using
        End Sub
    End Module
End Namespace
