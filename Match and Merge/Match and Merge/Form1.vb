Imports System.IO
Imports Microsoft.Office.Interop
Public Class Form1

    Public filepath1, filepath2 As String
    Public connection1, connection2 As ADODB.Connection
    Public datasource1, datasource2 As DataSet
    Public cat1, cat2 As ADOX.Catalog
    Public table1, table2 As ADOX.Table

    

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        OpenFileDialog1.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments : OpenFileDialog1.RestoreDirectory = False
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click, Button2.Click

        OpenFileDialog1.FileName = ""

        OpenFileDialog1.Filter = "Access File (*.mdb)|*.mdb"

        Dim feedback As DialogResult = OpenFileDialog1.ShowDialog
        If feedback = Windows.Forms.DialogResult.OK And OpenFileDialog1.FileName <> "" Then
            If sender.name = "Button1" Then
                TextBox1.Text = OpenFileDialog1.FileName
                cat1 = New ADOX.Catalog
                connection1 = New ADODB.Connection
                LoadAccessDB_ToComboBox(filepath1, cat1, connection1, ComboBox1)
            ElseIf sender.name = "Button2" Then
                TextBox2.Text = OpenFileDialog1.FileName
                cat2 = New ADOX.Catalog
                connection2 = New ADODB.Connection
                LoadAccessDB_ToComboBox(filepath2, cat2, connection2, ComboBox2)
            End If

        End If
    End Sub

    Dim DataTable1, DataTable2 As DataTable
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Me.Hide()
        Form2.Show()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged, ComboBox2.SelectedIndexChanged
        If ComboBox1.Items.Count > 0 And ComboBox2.Items.Count > 0 Then
            table1 = cat1.Tables.Item(ComboBox1.SelectedItem)
            table2 = cat2.Tables.Item(ComboBox2.SelectedItem)
            Button3.Enabled = True
        Else
            Button3.Enabled = False
        End If


    End Sub

    Private Sub Btn_clear_Click(sender As Object, e As EventArgs) Handles btn_clear.Click
        TextBox1.Text = "" : TextBox2.Text = ""
        If ComboBox1.Items.Count > 0 Then
            ComboBox1.Items.Clear() : ComboBox1.Text = ""
        End If
        If ComboBox2.Items.Count > 0 Then
            ComboBox2.Items.Clear() : ComboBox2.Text = ""
        End If
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged, TextBox2.TextChanged
        If sender.name = "TextBox1" Then
            filepath1 = sender.text
        Else
            filepath2 = sender.text
        End If
    End Sub

    Sub LoadAccessDB_ToComboBox(ByVal FileDirectory As String, ByRef cat As ADOX.Catalog, ByRef connection As ADODB.Connection, Combobox As ComboBox)
        Dim AccessObject As Access.Application = New Access.Application

        Dim connectionstring As String

        AccessObject.OpenCurrentDatabase(FileDirectory, False)
        connectionstring = AccessObject.CurrentProject.BaseConnectionString

        connection.ConnectionString = connectionstring.Replace("Microsoft.ACE.OLEDB.12.0", "Microsoft.Jet.OLEDB.4.0")

        'AccessObject.CloseCurrentDatabase()

        AccessObject.Quit()

        connection.Open()

        cat.ActiveConnection = connection

        Combobox.Items.Clear()

        If cat.Tables.Count > 0 Then
            For Each table As ADOX.Table In cat.Tables
                If table.Type = "TABLE" Then
                    Combobox.Items.Add(table.Name)
                End If
            Next
            Combobox.SelectedIndex = 0
        End If

    End Sub

   
End Class
