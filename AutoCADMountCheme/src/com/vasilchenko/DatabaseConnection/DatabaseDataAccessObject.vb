Imports AutoCADMountCheme.com.vasilchenko.classes
Imports AutoCADMountCheme.com.vasilchenko.Modules
Imports Autodesk.AutoCAD.ApplicationServices

Namespace com.vasilchenko.DatabaseConnection
    Module DatabaseDataAccessObject
        Friend Function GetAllElementTag() As List(Of String)
            Dim acElementsList As New List(Of String)
            Const sqlQuery As String =
                      "SELECT DISTINCT [COMPPINS].[TAGNAME] " &
                      "FROM [COMPPINS] " &
                      "INNER JOIN [COMP] ON ([COMPPINS].[TAGNAME] = [COMP].[TAGNAME] AND [COMPPINS].[HDL] = [COMP].[HDL]) " &
                      "WHERE ([COMPPINS].[TERM] IS NOT NULL " &
                      "AND ([COMPPINS].[TAGNAME] NOT IN (SELECT [TAGNAME] FROM [PNLCOMP] WHERE [BLOCK] LIKE 'WD*'))) " &
                      "ORDER BY [COMPPINS].[TAGNAME] ASC"
            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(sqlQuery)
            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    acElementsList.Add(objRow.item("TAGNAME"))
                Next
            End If
            Return acElementsList
        End Function

        Public Function GetSingleElementData(elementTag As String) As ElementClass
            Dim acCurElement As New ElementClass

            Dim sqlQuery As String =
                    "SELECT [COMPPINS].[TERM], [COMPPINS].[TERMCODE], [COMP].[PAR1_CHLD2], [COMP].[BLOCK], [COMP].[MFG], [COMPPINS].[HDL] " &
                    "FROM [COMPPINS] " &
                    "INNER JOIN [COMP] ON ([COMPPINS].[TAGNAME] = [COMP].[TAGNAME] AND [COMPPINS].[HDL] = [COMP].[HDL]) " &
                    "WHERE ([COMPPINS].[TERM] IS NOT NULL " &
                    "AND ([COMPPINS].[TAGNAME] =  '" & elementTag & "' AND [COMPPINS].[TERM] IS NOT NULL)) " &
                    "ORDER BY [COMPPINS].[TAGNAME] ASC, [COMP].[PAR1_CHLD2] ASC, [COMPPINS].[TERMCODE] ASC"

            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(sqlQuery)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim acCurTerm As New TermClass
                    Dim isPlc = objRow.item("BLOCK").ToString.ToUpper.Contains("PLCIO")
                    Dim intTermcode

                    If isPlc Then
                        If Not IsDBNull(objRow.item("MFG")) AndAlso objRow.item("MFG").Equals("ABB") Then
                            intTermcode = AdditionalFunctionsModule.GetLastNumericFromString(objRow.item("TERM"))
                        Else
                            intTermcode = AdditionalFunctionsModule.GetFirstNumericFromString(objRow.item("TERM"))
                        End If
                    Else
                        If IsNumeric(objRow.item("TERM")) Then
                            intTermcode = CDbl(objRow.item("TERM"))
                        Else
                            intTermcode = AdditionalFunctionsModule.GetFirstNumericFromString(objRow.item("TERMCODE"))
                        End If
                    End If

                    If acCurElement.TermList.Count <> 0 then
                        If _
                            Not isPlc AndAlso Not IsNumeric(objRow.item("TERM")) AndAlso
                            objRow.item("PAR1_CHLD2").Equals("2") Then
                            intTermcode = acCurElement.TermList.Values.Count + 1
                        End If
                        acCurTerm.Term = objRow.item("TERM")
                        acCurTerm.TermCode = intTermcode
                        acCurElement.TermList.Add(intTermcode, acCurTerm)
                    Else
                        acCurElement.TagName = elementTag
                        acCurElement.Handle = objRow.item("HDL")
                        acCurTerm.Term = objRow.item("TERM")
                        acCurTerm.TermCode = intTermcode
                        acCurElement.AddTerm(intTermcode, acCurTerm)
                    End If

                Next objRow
            End If
            Return acCurElement
        End Function

        Friend Function GetAllElementData(elementTag As String, acElementsList As List(Of  ElementClass))
            acElementsList.Add(GetSingleElementData(elementTag))
            Return acElementsList
        End Function

        Friend Sub FillElementData(element As ElementClass)

            Dim sqlQuery = "SELECT [MFG], [CAT] " &
                           "FROM [COMP] " &
                           "WHERE [TAGNAME] = '" & element.TagName & "' " &
                           "AND [MFG] IS NOT NULL AND [CAT] IS NOT NULL"

            Dim objDataTable As Data.DataTable = DatabaseConnections.GetOleDataTable(sqlQuery)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    If Not IsDBNull(objRow.item("MFG")) Then element.Manufacture = objRow.item("MFG") Else _
                        MsgBox("Елемент " & element.TagName & ". Не указано производителя", MsgBoxStyle.Information)
                    If Not IsDBNull(objRow.item("CAT")) Then element.CatalogName = objRow.item("CAT") Else _
                        MsgBox("Елемент " & element.TagName & ". Не указано марка", MsgBoxStyle.Information)

                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                        "Элемент " & element.CatalogName & " добавлен в список" & vbCrLf)
                Next objRow
            End If
        End Sub

        Friend Sub FillBlockPath(element As ElementClass)
            Dim sqlQuery = ""
            Try
                sqlQuery = "SELECT [BLKNAM] " &
                           "FROM [WD_MOUNT_GRAPHICS] " &
                           "WHERE [CATALOG] LIKE '" & element.CatalogName & "'"

                Dim objDataTable As Data.DataTable = DatabaseConnection.GetSqlDataTable(sqlQuery,
                                                                                        My.Settings.
                                                                                           footprint_lookupSQLConnectionString)
                If Not IsNothing(objDataTable) Then _
                    element.WdBlockPath = My.Settings.MOUNT_GRAPHICS_PATH & objDataTable.Rows(0).Item("BLKNAM")
            Catch ex As Exception
                MsgBox("Объект не найден в графической базе. Запрос: " & sqlQuery)
            End Try
        End Sub

        Friend Sub FillElementConnections(element As ElementClass)

            For Each connection In element.TermList.Values
                FillTermConnections(connection, element.TagName)
            Next
        End Sub

        Private Sub FillTermConnections(connection As TermClass, tagname As String)

            Dim sqlQuery =
                    "SELECT [NAM1], [PIN1], [NAM2], [PIN2], [CBL], [TERMDESC1], [TERMDESC2], [NAMHDL1], [NAMHDL2], [WLAY1], [WLAY2] " &
                    "FROM WFRM2ALL " &
                    "WHERE (([NAM1] = '" & tagname & "' AND [PIN1] = '" & connection.Term & "') OR ([NAM2]= '" & tagname &
                    "' AND [PIN2]= '" & connection.Term & "'))"

            Dim objDataTable As Data.DataTable = DatabaseConnection.GetOleDataTable(sqlQuery)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim pin = New PinClass()
                    If objRow.Item("NAM1").Equals(tagname) AndAlso objRow.Item("PIN1").Equals(connection.Term) Then
                        'IIf (Not IsNothing(objRow.Item("TERMDESC2")), TruePart:=pin = objRow.Item("TERMDESC2"), FalsePart:=pin = objRow.Item("PIN2"))
                        If Not IsDBNull(objRow.Item("NAM2")) Then pin.TagName = objRow.Item("NAM2")
                        If Not IsDBNull(objRow.Item("PIN2")) Then pin.Pin = objRow.Item("PIN2")
                        If Not IsDBNull(objRow.Item("WLAY2")) Then _
                            connection.WireType = RecognizeWireType(objRow.Item("WLAY2"))
                    Else
                        If Not IsDBNull(objRow.Item("NAM1")) Then pin.TagName = objRow.Item("NAM1")
                        If Not IsDBNull(objRow.Item("PIN1")) Then pin.Pin = objRow.Item("PIN1")
                        If Not IsDBNull(objRow.Item("WLAY1")) Then _
                            connection.WireType = RecognizeWireType(objRow.Item("WLAY1"))
                    End If

                    If Not IsDBNull(objRow.Item("CBL")) Then pin.CableName = objRow.Item("CBL")

                    connection.AddConnection(pin)
                Next objRow
            End If
        End Sub

        Private Function RecognizeWireType(wireName As String) As String
            wireName = wireName.Replace(",", ".")
            Dim section = AdditionalFunctionsModule.GetFirstNumericFromString(wireName)
            Select Case section
                Case 5
                    Return "***"
                Case 10
                    Return "*"
                Case 25
                    Return "**"
                Case Else
                    Return ""
            End Select
        End Function

        Friend Function GetAddressMarkingList() As List(Of MarkingClass)
            Dim markingList As New List(Of MarkingClass)

            Const sqlQuery As String = "SELECT [NAM1], [PIN1], [NAM2], [PIN2], [CBL], [TERMDESC1], [TERMDESC2] " &
                                       "FROM WFRM2ALL " &
                                       "WHERE [NAM1] IS NOT NULL AND [PIN1] IS NOT NULL AND [NAM2] IS NOT NULL AND [PIN2] IS NOT NULL " &
                                       "ORDER BY [NAM1], [NAM2]"

            Dim objDataTable As Data.DataTable = DatabaseConnection.GetOleDataTable(sqlQuery)

            If Not IsNothing(objDataTable) Then
                For Each objRow In objDataTable.Rows
                    Dim pinOne = New PinClass(objRow.Item("NAM1"), objRow.Item("PIN1"))
                    Dim pinTwo = New PinClass(objRow.Item("NAM2"), objRow.Item("PIN2"))
                    If Not IsDBNull(objRow.Item("CBL")) Then
                        pinOne.CableName = objRow.Item("CBL")
                        pinTwo.CableName = objRow.Item("CBL")
                    End If
                    markingList.Add(New MarkingClass(pinOne, pinTwo))
                Next objRow
            End If

            Return markingList
        End Function
    End Module
End Namespace