Imports System.Windows.Forms
Imports AutoCADMountCheme.com.vasilchenko.Classes
Imports AutoCADMountCheme.com.vasilchenko.DatabaseConnection

Public Class ElementSelector

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        Hide()
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        Close()
    End Sub

    Private Sub ElementSelector_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Const strSqlQuery As String = "SELECT DISTINCT [COMPPINS].[TAGNAME] " &
                                        "FROM [COMPPINS] " &
                                        "WHERE [COMPPINS].[TERM] IS NOT NULL " &
                                        "ORDER BY [COMPPINS].[TAGNAME] ASC"
        Dim objDataTable As System.Data.DataTable = DatabaseConnections.GetOleDataTable(strSqlQuery)
        Dim acDrawingList As New SortedList(Of Integer, String)
        Dim i As Short = 0
        If Not IsNothing(objDataTable) Then
            For Each objRow In objDataTable.Rows
                If Not IsDBNull(objRow.item("TAGNAME")) Then
                    acDrawingList.Add(i, objRow.item("TAGNAME"))
                    i += 1
                End If
            Next objRow
        End If

        i=0

        'lvSheets.Columns.Add("Checked", Windows.HorizontalAlignment.Left)

        lvSheets.MultiSelect = True
        lvSheets.FullRowSelect = True
        lvSheets.Sorting = SortOrder.Ascending
        lvSheets.GridLines = True
        lvSheets.View = View.Details
        lvSheets.AutoSize = True
        AutoSize = True

        lvSheets.Sorting = SortOrder.Ascending
        'lvSheets.Sort()

        lvSheets.Columns.Add("##", 50, HorizontalAlignment.Left)
        lvSheets.Columns.Add("TagName", 120, HorizontalAlignment.Left)
        'lvSheets.CheckBoxes = True
        Dim arrLvItem(acDrawingList.Count - 1) As Windows.Forms.ListViewItem

        For Each pair As KeyValuePair(Of Integer, String) In acDrawingList
            arrLvItem(i) = New Windows.Forms.ListViewItem(pair.Key)
            arrLvItem(i).SubItems.Add(pair.Value)
            i += 1
        Next pair
        lvSheets.Items.AddRange(arrLvItem)
    End Sub

    Private _mSortingColumn As ColumnHeader

    Private Sub ListView1_ColumnClick(ByVal sender As Object, ByVal e As System.Windows.Forms.ColumnClickEventArgs) Handles lvSheets.ColumnClick
        ' Get the new sorting column. 
        Dim newSortingColumn As ColumnHeader = lvSheets.Columns(e.Column)
        ' Figure out the new sorting order. 
        Dim sortOrder As SortOrder
        If _mSortingColumn Is Nothing Then
            ' New column. Sort ascending. 
            sortOrder = SortOrder.Ascending
        Else ' See if this is the same column. 
            If newSortingColumn.Equals(_mSortingColumn) Then
                ' Same column. Switch the sort order. 
                If _mSortingColumn.Text.StartsWith("> ") Then
                    sortOrder = SortOrder.Descending
                Else
                    sortOrder = SortOrder.Ascending
                End If
            Else
                ' New column. Sort ascending. 
                sortOrder = SortOrder.Ascending
            End If
            ' Remove the old sort indicator. 
            _mSortingColumn.Text = _mSortingColumn.Text.Substring(2)
        End If
        ' Display the new sort order. 
        _mSortingColumn = newSortingColumn
        If sortOrder = SortOrder.Ascending Then
            _mSortingColumn.Text = "> " & _mSortingColumn.Text
        Else
            _mSortingColumn.Text = "< " & _mSortingColumn.Text
        End If
        ' Create a comparer. 
        lvSheets.ListViewItemSorter = New ListViewColumnSorter(e.Column, sortOrder)
        ' Sort. 
        lvSheets.Sort()
    End Sub
    
End Class