Imports Autodesk.AutoCAD.Geometry

Namespace com.vasilchenko.classes

    Public Class TermClass

        Public Property Term As String
        Public Property TermCode As String
        Public Property Direction As Enums.Direction
        Public Property WireType As String
        Public Property BasePoint As Point3d
        Private _connectionsList As List(Of PinClass)

        Public Function ConnectionsList() As List(Of PinClass)
            Return _connectionsList
        End Function

        Public Function GetConnections(index As Integer) As String
            Return _connectionsList(index).TagName & ":" & _connectionsList(index).Pin
        End Function

        Public Function HasCableInConnection(index As Integer) As String
            Return Not IsNothing(_connectionsList(index).CableName)
        End Function

        Public Function GetCableInConnection(index As Integer) As String
            Return IIf(IsNothing(_connectionsList(index).CableName), Nothing, _connectionsList(index).CableName)
        End Function

        Public Sub AddConnection(con As PinClass)
            _connectionsList.Add(con)
        End Sub

        Public Sub New()
            _connectionsList = New List(Of PinClass)
        End Sub

        Protected Overrides Sub Finalize()
            _connectionsList = Nothing
        End Sub

    End Class

End Namespace

