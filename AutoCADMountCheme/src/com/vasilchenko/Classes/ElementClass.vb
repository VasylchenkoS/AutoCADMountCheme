Namespace com.vasilchenko.classes

    Public Class ElementClass
        
        Public Property TagName As String
        Private Property _termList as SortedList(Of Integer, TermClass)
        Public Property Manufacture As String
        Public Property CatalogName As String
        Public Property WdBlockPath As String
        Public Property Handle As String
        
        Public Sub AddTerm(TermCode As String, Term As TermClass)
            _termList.add(CInt(TermCode), Term)
        End Sub

        Public ReadOnly Property TermList As SortedList(Of Integer, TermClass)
            Get
                Return _termList
            End Get
        End Property
        
        Public Sub New()
            _termList = new SortedList(Of Integer, TermClass)
        End Sub

        Protected Overrides Sub Finalize()
            _termList = Nothing
        End Sub

    End Class

End Namespace

