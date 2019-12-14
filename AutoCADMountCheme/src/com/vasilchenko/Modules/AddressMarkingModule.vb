Imports System.IO
Imports AutoCADMountCheme.com.vasilchenko.classes
Imports AutoCADMountCheme.com.vasilchenko.DatabaseConnection
Imports Autodesk.AutoCAD.ApplicationServices
Imports Autodesk.Electrical.Project

Namespace com.vasilchenko.Modules
    Module AddressMarkingModule
        Friend Sub CreateFileWithAddressMarking()
            Dim acAddrList As List(Of MarkingClass) = DatabaseDataAccessObject.GetAddressMarkingList()

            'Dim path = "D:\Васильченко\1. Проекты в разработке\Днепрооблэнерго\4. СЕС Солар Фарм-3\АСКУЭ\Мар'янська\Конструкторская документация\EU114SOE\"
            Dim ioFilePath = Path.GetDirectoryName(Application.DocumentManager.MdiActiveDocument.Name)
            Dim ioFileName = ProjectManager.GetInstance().GetActiveProject().GetProjectID
            ioFileName = ioFileName.Substring(ioFileName.LastIndexOf("\", StringComparison.Ordinal))
            ioFileName = ioFileName.Substring(0,
                                              ioFileName.Length -
                                              (ioFileName.Length - ioFileName.LastIndexOf(".", StringComparison.Ordinal))) & "_Mark.txt"
'            ioFileName = "\EU114SOE_Mark.txt"

            If Not File.Exists(ioFilePath & ioFileName) Then
                Dim fs As FileStream = File.Create(ioFilePath & ioFileName)
                fs.Close()
            Else
                My.Computer.FileSystem.WriteAllText(ioFilePath & ioFileName, "", False)
            End If
            Using objStreamWriter As New StreamWriter(ioFilePath & ioFileName)
                For i = 0 To acAddrList.Count - 1
                    objStreamWriter.WriteLine(acAddrList(i).PinOne.TagName & ":" & acAddrList(i).PinOne.Pin & "/" & acAddrList(i).PinTwo.TagName & ":" & acAddrList(i).PinTwo.Pin)
                    objStreamWriter.WriteLine(acAddrList(i).PinTwo.TagName & ":" & acAddrList(i).PinTwo.Pin & "/" & acAddrList(i).PinOne.TagName & ":" & acAddrList(i).PinOne.Pin)
                Next
            End Using
        End Sub
    End Module
End Namespace
