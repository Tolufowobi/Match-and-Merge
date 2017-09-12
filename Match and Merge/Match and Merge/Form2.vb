Public Class Form2


    Dim Table1, Table2 As ADOX.Table
    Dim RecordSet1, RecordSet2, NewRecordSet As ADODB.Recordset

    Dim Table1SelectedColumns, Table2SelectedColumns As String
    Dim PotentialNewColumns As DataColumnCollection
    Dim TableNumbertags, typetags As Collection

    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        MergeColumnIndexes1 = New Collection
        Mergecolumnindexes2 = New Collection

        Table1 = Form1.table1 : Table2 = Form1.table2
        lbl_Table1.Text = Table1.Name
        Lbl_table2.Text = Table2.Name

        'object initialization

        TableNumbertags = New Collection : typetags = New Collection

        For Each clm As ADOX.Column In Table1.Columns
            cmbbx_columns1.Items.Add(clm.Name)
            chkLbx_columns.Items.Add(clm.Name)
            TableNumbertags.Add(1, clm.Name)
            typetags.Add(clm.Type, clm.Name)
        Next
        cmbbx_columns1.SelectedIndex = 0

        For Each clm As ADOX.Column In Form1.table2.Columns
            cmbbx_columns2.Items.Add(clm.Name)

            If TableNumbertags.Contains(clm.Name) Then
                MsgBox("Column '" & clm.Name & "' in table '" & Table2.Name & "' duplicates a column of the same name in '" & Table2.Name & "' and wouldn't be added.", , "Merge and Match")
            Else
                chkLbx_columns.Items.Add(clm.Name)
                TableNumbertags.Add(2, clm.Name)
                typetags.Add(clm.Type, clm.Name)
            End If
        Next
        cmbbx_columns2.SelectedIndex = 0


        cmbbx_compare.SelectedIndex = 0

    End Sub

    Dim MergeColumnIndexes1, Mergecolumnindexes2 As Collection
    Dim index1, index2 As Integer

    Dim BooleanFlag As Integer = -1
    Dim ComparatorCollection As Collection = New Collection

    Private Sub cmbbx_columns1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbbx_columns1.SelectedIndexChanged, cmbbx_columns2.SelectedIndexChanged
        If sender.items.count <> 0 Then ' if the number of items in the ccombobox isn't zero

            If sender.name = "cmbbx_columns1" Then ' if the column name was selected from the first combobox
                index1 = sender.selectedindex
            ElseIf sender.name = "cmbbx_columns2" Then 'if the column name was selected from the second combobox
                index2 = sender.selectedindex
            End If

            ValidateOperatorOptions(sender.selecteditem) ' call method to check if datatype of selected column can work with the In operator (Only a string datatype can work with the iIn operator)

        End If
    End Sub

    Dim comparecolumns1 As Collection = New Collection
    Dim comparecolumns2 As Collection = New Collection
    Private Sub Btn_add_Click(sender As Object, e As EventArgs) Handles btn_Add.Click

        'If MergeColumnIndexes1.Contains(cmbbx_columns1.SelectedItem) = False Then
        '    MergeColumnIndexes1.Add(cmbbx_columns1.SelectedItem, cmbbx_columns1.SelectedItem)
        'End If
        'If Mergecolumnindexes2.Contains(cmbbx_columns2.SelectedItem) = False Then
        '    Mergecolumnindexes2.Add(cmbbx_columns2.SelectedItem, cmbbx_columns2.SelectedItem)
        'End If

        MergeColumnIndexes1.Add(cmbbx_columns1.SelectedItem)
        Mergecolumnindexes2.Add(cmbbx_columns2.SelectedItem)

        ComparatorCollection.Add(cmbbx_compare.SelectedIndex)
        Lbox_CompareColumns.Items.Add(cmbbx_columns1.SelectedItem & " " & cmbbx_compare.SelectedItem & " " & cmbbx_columns2.SelectedItem)
        Lbox_CompareColumns.SelectedIndex = Lbox_CompareColumns.TopIndex

        '  comparecolumns1.Add(cmbbx_columns1.SelectedItem, cmbbx_columns1.SelectedItem) : comparecolumns2.Add(cmbbx_columns2.SelectedItem, cmbbx_columns2.SelectedItem)
        If Lbox_CompareColumns.Items.Count > 1 And Panel1.Visible = False Then
            Panel1.Visible = True
            BooleanFlag = 0
        End If

    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged, RadioButton2.CheckedChanged
        If sender.name = "RadioButton1" Then ' for Or
            BooleanFlag = 0
        ElseIf sender.name = "RadioButton2" Then ' For And
            BooleanFlag = 1
        End If
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles Lbox_CompareColumns.SelectedIndexChanged
        If sender.items.count = 0 Then
            btn_remove.Enabled = False
            chkLbx_columns.Visible = False
        Else
            btn_remove.Enabled = True
            chkLbx_columns.Visible = True
        End If


    End Sub

    Private Sub btn_remove_Click(sender As Object, e As EventArgs) Handles btn_remove.Click
        MergeColumnIndexes1.Remove(Lbox_CompareColumns.SelectedIndex + 1) : Mergecolumnindexes2.Remove(Lbox_CompareColumns.SelectedIndex + 1)
        ComparatorCollection.Remove(Lbox_CompareColumns.SelectedIndex + 1)
        Lbox_CompareColumns.Items.RemoveAt(Lbox_CompareColumns.SelectedIndex)

        If Lbox_CompareColumns.Items.Count < 2 Then
            Panel1.Visible = False
            BooleanFlag = -1
        End If

        If Lbox_CompareColumns.Items.Count > 0 Then
            Lbox_CompareColumns.SelectedIndex = Lbox_CompareColumns.TopIndex
        End If
    End Sub

    Private Sub chkLbx_columns_itemcheck(sender As Object, e As EventArgs) Handles chkLbx_columns.ItemCheck
        If chkLbx_columns.SelectedIndices.Count > 0 Then
            btn_merge.Enabled = True
        Else : btn_merge.Enabled = False
        End If
    End Sub

    Dim table1Rowi As Integer : Dim table2rowi As Integer = 0
    Dim CompareResultsCount As Integer
    Dim ResultsStore() As Boolean
    Dim datatable1, datatable2 As DataTable
    Dim NewTable As ADOX.Table
    Private Sub btn_merge_Click(sender As Object, e As EventArgs) Handles btn_merge.Click

        Dim TableHAsBeenAdded As Boolean

        NewTable = New ADOX.Table With
                                        {.Name = "Merge Table", .ParentCatalog = Form1.cat1}
        Try

            Dim sql1 As String = "SELECT * FROM [" & Table1.Name & "]"
            Dim sql2 As String = "SELECT * FROM [" & Table2.Name & "]"

            Dim connection1 As OleDb.OleDbConnection = New OleDb.OleDbConnection(Form1.connection1.ConnectionString)
            Dim Connection2 As OleDb.OleDbConnection = New OleDb.OleDbConnection(Form1.connection2.ConnectionString)

            Dim command1 As OleDb.OleDbCommand = New OleDb.OleDbCommand(sql1, connection1)
            Dim command2 As OleDb.OleDbCommand = New OleDb.OleDbCommand(sql2, Connection2)

            Dim adapter1 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(command1)
            Dim adapter2 As OleDb.OleDbDataAdapter = New OleDb.OleDbDataAdapter(command2)

            datatable1 = New DataTable(Table1.Name)
            datatable2 = New DataTable(Table2.Name)

            adapter1.Fill(datatable1) : adapter2.Fill(datatable2)

            ConvertNulls(datatable1)
            ConvertNulls(datatable2)

            If datatable1.Rows.Count > 0 And datatable2.Rows.Count > 0 Then
                'add all the selected columns to the newly created table
                For i As Integer = 0 To chkLbx_columns.CheckedItems.Count - 1

                    Dim newcolumn As ADOX.Column = New ADOX.Column With {.Name = chkLbx_columns.CheckedItems.Item(i), .Type = typetags(.Name)}
                    NewTable.Columns.Append(newcolumn)
                    'simultaneously add the selected items as columns to the datagridview object
                    TableMonitor.DGV.Columns.Add(newcolumn.Name, newcolumn.Name)
                Next

                ' add the table to the containing catalog of the fisrt merge table
                Form1.cat1.Tables.Append(NewTable)

                TableHAsBeenAdded = True ' for exception handling purposes

                'Dim sql3 As String = "SELECT * FROM [" & NewTable.Name & "]"
                'NewRecordSet = New ADODB.Recordset
                'NewRecordSet.Open(sql3, Form1.cat1.ActiveConnection, ADODB.CursorTypeEnum.adOpenDynamic, ADODB.LockTypeEnum.adLockBatchOptimistic)

                TableMonitor.Show()
                Me.Hide()

                While table1Rowi < datatable1.Rows.Count ' 0 while the row index  hasnt gone beyond the last row in Table 1

                    While table2rowi < datatable2.Rows.Count ' while the second row index hasnt gone beyond the last row in Table2

                        CompareResultsCount = Lbox_CompareColumns.Items.Count

                        ReDim ResultsStore(CompareResultsCount - 1)

                        MergeAndWriteCentre(table1Rowi, table2rowi, BooleanFlag)

                        table2rowi += 1

                    End While

                    table2rowi = 0 ' reset the table 2 index

                    table1Rowi += 1

                End While

            Else

                MsgBox("one or both tables seleted is empty." & vbNewLine & "Please make sure both tables contain records.", MsgBoxStyle.Information, "Match and Merge")
                Exit Sub

            End If

        Catch ex As Exception

            MsgBox(ex.Message & vbNewLine & " Table 1 Row index: " & table1Rowi & vbNewLine & "Table 2 Row index: " & table2rowi)

            If TableHAsBeenAdded = True Then
                NewRecordSet.Delete(ADODB.AffectEnum.adAffectAll)
                Form1.cat1.Tables.Delete(NewTable.Name)
                table1Rowi = 0 : table2rowi = 0
            End If

        Finally
            Dim AccessApplication As Microsoft.Office.Interop.Access.Application = New Microsoft.Office.Interop.Access.Application

            AccessApplication.OpenCurrentDatabase(Form1.filepath1)

            AccessApplication.DoCmd.OpenTable(NewTable.Name, Microsoft.Office.Interop.Access.AcView.acViewNormal, Microsoft.Office.Interop.Access.AcOpenDataMode.acEdit)

            AccessApplication.Application.Visible = True

            Form1.connection1.Close() : Form1.connection2.Close()
            Beep()
            MsgBox("DONE!")

        End Try


    End Sub

    Sub MergeAndWriteCentre(ByVal Rowindex1 As Integer, ByVal Rowindex2 As Integer, ByVal BooleanOperatorIndex As Integer)
        'run a comparism on between all columns in both comparism collections

        For i = 1 To CompareResultsCount
            Dim data1 = datatable1.Rows(Rowindex1)(MergeColumnIndexes1(i)) : Dim data2 = datatable2.Rows(Rowindex2)(Mergecolumnindexes2(i))
            ResultsStore(i - 1) = ColumnsDataCompare(data1, data2, ComparatorCollection(i))
        Next

        Dim IsConditionSatisfied As Boolean = False

        ' decide when to write the necessary data to storage based on the boolean setting
        If BooleanOperatorIndex = -1 Then ' if only one pair of columns comparism is being executed
            '0
            If ResultsStore(0) = True Then ' if a match was made 
                IsConditionSatisfied = True
                GoTo WritingStation

            End If

        ElseIf BooleanOperatorIndex = 0 Then ' if the Or radio button was selected
            '0
            For i = 0 To ResultsStore.Length - 1

                If ResultsStore(i) = True Then
                    IsConditionSatisfied = True
                    Exit For
                End If
            Next

            GoTo WritingStation

        ElseIf BooleanOperatorIndex = 1 Then ' if the And radio button was selected

            IsConditionSatisfied = True

            For i = 0 To ResultsStore.Length - 1
                If ResultsStore(i) = False Then
                    IsConditionSatisfied = False
                    Exit For
                End If
            Next
            GoTo WritingStation

            '0
        End If

WritingStation:
        If IsConditionSatisfied = True Then
            RecordWrite(Rowindex1, Rowindex2)
        End If
    End Sub

    Function ColumnsDataCompare(ByVal value1 As Object, ByVal value2 As Object, ByVal comparatorindex As Integer) As Boolean
        Dim MatchMade As Boolean = False
        If comparatorindex = 0 Then 'the IN comparator was selected

            Dim stringcollection As List(Of String) = New List(Of String)
            stringcollection.AddRange(value2.Split({" "}, StringSplitOptions.RemoveEmptyEntries)) ' break the string in the second variable into seperate words _ 
            'using an blank space delimiter

            If stringcollection.Count > 1 Then ' if theres more than one word in the resultant collection

                For i = 0 To stringcollection.Count - 1 'compare the value in the first variable with each word in the collection iteratively

                    If i <> stringcollection.Count - 1 Then ' if the loop hasn't gotten to the last element in the collection, include present and future element concatenation _ 
                        'in the comparism

                        If value1 = stringcollection(i) Or value1 = stringcollection(i) & " " & stringcollection(i + 1) Then

                            If value1 <> " " And value1 <> Nothing Then

                                MatchMade = True
                                Exit For
                            End If

                        End If

                    Else ' if the loop is at the last string element in the collection, compare without concatenation
                        If value1 = stringcollection(i) Then
                            If value1 <> " " And value1 <> Nothing Then
                                MatchMade = True
                                Exit For
                            End If
                        End If

                    End If

                Next


            Else ' if the collection isn't greater than one word, incase of an end of string ommission, check if the first value is a substring of the second

                If value2.Contains(value1) = True And value1 <> "" Then
                    MatchMade = True
                Else : MatchMade = False
                End If

            End If

        ElseIf comparatorindex = 1 Then ' the EQUAL-to comparator was selected
            If value1 = value2 Then
                MatchMade = True
            Else : MatchMade = False
            End If

        Else : MatchMade = False ' if the comparator index isn't 0 or 1-- unattainable

        End If

        Return MatchMade

    End Function

    Sub RecordWrite(ByVal index1 As Integer, ByVal index2 As Integer)

        Dim vals(chkLbx_columns.CheckedItems.Count - 1) As String
        Dim InsertString As String = "INSERT INTO [" & NewTable.Name & "] VALUES( "
        Dim values As String = ""
        Dim loopindex As Integer = 0

        Try

            For Each i As String In chkLbx_columns.CheckedItems ' run through all the checked checkboxes in the checkbox collection

                If TableNumbertags(i) = 1 Then ' if the column originally belongs to the first table, using the name of the checkbox to as a reference key to the collection item

                    values &= "'" & datatable1.Rows.Item(index1)(i) & "',"
                    vals(loopindex) = datatable1.Rows.Item(index1)(i)

                ElseIf TableNumbertags(i) = 2 Then 'if the column originally belongs to the second table, ***

                    values &= "'" & datatable2.Rows.Item(index2)(i) & "',"
                    vals(loopindex) = datatable2.Rows.Item(index2)(i)

                End If
                loopindex += 1
            Next

            values = values.Remove(values.Length - 1, 1)
            values = values & ")"

            InsertString = InsertString & values

            Dim connection As OleDb.OleDbConnection = New OleDb.OleDbConnection(Form1.connection1.ConnectionString)
            Dim command As OleDb.OleDbCommand = New OleDb.OleDbCommand(InsertString, connection)

            connection.Open()

            Dim feedback As Integer = command.ExecuteNonQuery()

            If feedback > 0 Then
                TableMonitor.DGV.Rows.Add(vals)
                TableMonitor.DGV.Rows.Item(TableMonitor.DGV.Rows.GetLastRow(DataGridViewElementStates.None)).Selected = True
                TableMonitor.DGV.Refresh()
            End If

            connection.Close()
            connection.Dispose()
            command = Nothing

            '  NewRecordSet.AddNew(fields, vals)
            '  NewRecordSet.Update()
        Catch ex As Exception
            'Dim fieldsToValues As String = ""
            'If fields.Count > 0 Then
            '    For Each i In fields
            '        fieldsToValues = fieldsToValues & fields(i) & "-" & vals(i) & vbNewLine
            '    Next
            '    MsgBox(ex.Message & vbNewLine & fieldsToValues, , "Write Centre")
            'Else
            MsgBox(ex.Message, , "Write Centre")
            'End If
        End Try
    End Sub

    Sub ValidateOperatorOptions(ColumnName As String)
        If typetags(ColumnName) = ADOX.DataTypeEnum.adDouble Or typetags(ColumnName) = ADOX.DataTypeEnum.adInteger Or typetags(ColumnName) = ADOX.DataTypeEnum.adSingle Then
            'if the datatype of the column corresponds to a string, disable the operator options combobox and selec the EQUALS operator automatically
            cmbbx_compare.SelectedIndex = 1
            cmbbx_compare.Enabled = False
        Else
            cmbbx_compare.SelectedIndex = 0
            cmbbx_compare.Enabled = True
        End If
    End Sub

    Sub ConvertNulls(ByRef datatable As DataTable)
        Try
            Dim row As DataRow
            For Each row In datatable.Rows
                Dim i As Integer = 0
                While i < datatable.Columns.Count

                    If IsDBNull(row(i)) Then
                        If datatable.Columns(i).DataType.Name = "Integer" Or datatable.Columns(i).DataType.Name = "Double" Then
                            row(i) = 0
                        ElseIf datatable.Columns(i).DataType.Name = "String" Then
                            row(i) = ""
                        End If
                    End If
                    i += 1
                End While

            Next
        Catch ex As Exception
            MsgBox(ex.Message, MsgBoxStyle.Exclamation, "Data Services Error")
        End Try
    End Sub

    
End Class