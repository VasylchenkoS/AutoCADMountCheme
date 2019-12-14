' (C) Copyright 2011 by  
'
Imports Autodesk.AutoCAD.Runtime
Imports Autodesk.AutoCAD.ApplicationServices
Imports AutoCADMountCheme.com.vasilchenko.Modules

'' This line is not mandatory, but improves loading performances
'<Assembly: CommandClass(GetType(AutoCADMountCheme.Commands))>

Namespace com.vasilchenko

    Public Class Commands

        <CommandMethod("ASU_Mount_Chema_Builder", CommandFlags.Session)>
        Public Shared Sub Builder()

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            If Application.GetSystemVariable("MIRRTEXT") = "1" Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("MIRRTEXT variable set to 0")
                Application.SetSystemVariable("MIRRTEXT", 0)
            End If

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()

                Try
                    ElementListMakerModule.BuildMountSchema()
                Catch ex As System.Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using

        End Sub

        <CommandMethod("ASU_Mount_Chema_Selector_Builder", CommandFlags.Session)>
        Public Shared Sub SelectorBuilder()

            Application.AcadApplication.ActiveDocument.SendCommand("(command ""_-Purge"")(command ""_ALL"")(command ""*"")(command ""_N"")" & vbCr)
            Application.AcadApplication.ActiveDocument.SendCommand("AEREBUILDDB" & vbCr)

            If Application.GetSystemVariable("MIRRTEXT") = "1" Then
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("MIRRTEXT variable set to 0")
                Application.SetSystemVariable("MIRRTEXT", 0)
            End If

            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()

                Try
                    Dim tagList As New List(Of String)


                    Dim ufElementSelector As New ElementSelector
                    ufElementSelector.ShowDialog()
                    For Each item As Object In ufElementSelector.lvSheets.SelectedItems
                        tagList.Add(item.SubItems.Item(1).Text)
                    Next

                    SingleElementListMakerModule.BuildMountElements(tagList)

                Catch ex As System.Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using

        End Sub

        <CommandMethod("ASU_Address_Marking", CommandFlags.Session)>
        Public Shared Sub AddressMarking()
            AddressMarkingModule.CreateFileWithAddressMarking()
        End Sub

        <CommandMethod("ASU_Mount_Redraw_Element", CommandFlags.Session)>
        Public Shared Sub MountRedraw()
            Using docLock As DocumentLock = Application.DocumentManager.MdiActiveDocument.LockDocument()
                Try
                    RedrawElementModule.RedrawElement()
                Catch ex As System.Exception
                    MsgBox("ERROR:[" & ex.Message & "]" & vbCr & "TargetSite: " & ex.TargetSite.ToString & vbCr & "StackTrace: " & ex.StackTrace, vbCritical, "ERROR!")
                End Try
            End Using
        End Sub
    End Class

End Namespace