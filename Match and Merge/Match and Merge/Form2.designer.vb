<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form2
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.Lbox_CompareColumns = New System.Windows.Forms.ListBox()
        Me.cmbbx_columns1 = New System.Windows.Forms.ComboBox()
        Me.cmbbx_columns2 = New System.Windows.Forms.ComboBox()
        Me.btn_Add = New System.Windows.Forms.Button()
        Me.btn_merge = New System.Windows.Forms.Button()
        Me.btn_remove = New System.Windows.Forms.Button()
        Me.chkLbx_columns = New System.Windows.Forms.CheckedListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.lbl_Table1 = New System.Windows.Forms.Label()
        Me.Lbl_table2 = New System.Windows.Forms.Label()
        Me.RadioButton1 = New System.Windows.Forms.RadioButton()
        Me.RadioButton2 = New System.Windows.Forms.RadioButton()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.cmbbx_compare = New System.Windows.Forms.ComboBox()
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Lbox_CompareColumns
        '
        Me.Lbox_CompareColumns.FormattingEnabled = True
        Me.Lbox_CompareColumns.Location = New System.Drawing.Point(25, 220)
        Me.Lbox_CompareColumns.Name = "Lbox_CompareColumns"
        Me.Lbox_CompareColumns.Size = New System.Drawing.Size(310, 95)
        Me.Lbox_CompareColumns.TabIndex = 0
        '
        'cmbbx_columns1
        '
        Me.cmbbx_columns1.FormattingEnabled = True
        Me.cmbbx_columns1.Location = New System.Drawing.Point(25, 69)
        Me.cmbbx_columns1.Name = "cmbbx_columns1"
        Me.cmbbx_columns1.Size = New System.Drawing.Size(154, 21)
        Me.cmbbx_columns1.TabIndex = 1
        '
        'cmbbx_columns2
        '
        Me.cmbbx_columns2.FormattingEnabled = True
        Me.cmbbx_columns2.Location = New System.Drawing.Point(181, 69)
        Me.cmbbx_columns2.Name = "cmbbx_columns2"
        Me.cmbbx_columns2.Size = New System.Drawing.Size(154, 21)
        Me.cmbbx_columns2.TabIndex = 2
        '
        'btn_Add
        '
        Me.btn_Add.Location = New System.Drawing.Point(288, 92)
        Me.btn_Add.Name = "btn_Add"
        Me.btn_Add.Size = New System.Drawing.Size(47, 23)
        Me.btn_Add.TabIndex = 4
        Me.btn_Add.Text = "Add"
        Me.btn_Add.UseVisualStyleBackColor = True
        '
        'btn_merge
        '
        Me.btn_merge.Enabled = False
        Me.btn_merge.Location = New System.Drawing.Point(288, 473)
        Me.btn_merge.Name = "btn_merge"
        Me.btn_merge.Size = New System.Drawing.Size(47, 23)
        Me.btn_merge.TabIndex = 5
        Me.btn_merge.Text = "Merge"
        Me.btn_merge.UseVisualStyleBackColor = True
        '
        'btn_remove
        '
        Me.btn_remove.Enabled = False
        Me.btn_remove.Location = New System.Drawing.Point(279, 321)
        Me.btn_remove.Name = "btn_remove"
        Me.btn_remove.Size = New System.Drawing.Size(56, 23)
        Me.btn_remove.TabIndex = 6
        Me.btn_remove.Text = "Remove"
        Me.btn_remove.UseVisualStyleBackColor = True
        '
        'chkLbx_columns
        '
        Me.chkLbx_columns.CheckOnClick = True
        Me.chkLbx_columns.FormattingEnabled = True
        Me.chkLbx_columns.Location = New System.Drawing.Point(25, 357)
        Me.chkLbx_columns.Name = "chkLbx_columns"
        Me.chkLbx_columns.Size = New System.Drawing.Size(220, 139)
        Me.chkLbx_columns.TabIndex = 7
        Me.chkLbx_columns.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(29, 53)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 13)
        Me.Label1.TabIndex = 8
        Me.Label1.Text = "Select a column"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(185, 53)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(129, 13)
        Me.Label2.TabIndex = 9
        Me.Label2.Text = "Select matching column 2"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(28, 204)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(93, 13)
        Me.Label3.TabIndex = 10
        Me.Label3.Text = "Merge Parameters"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(29, 341)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(181, 13)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "Select Columns for your Merge Table"
        '
        'lbl_Table1
        '
        Me.lbl_Table1.AutoSize = True
        Me.lbl_Table1.Location = New System.Drawing.Point(29, 22)
        Me.lbl_Table1.Name = "lbl_Table1"
        Me.lbl_Table1.Size = New System.Drawing.Size(43, 13)
        Me.lbl_Table1.TabIndex = 12
        Me.lbl_Table1.Text = "Table 1"
        '
        'Lbl_table2
        '
        Me.Lbl_table2.AutoSize = True
        Me.Lbl_table2.Location = New System.Drawing.Point(185, 22)
        Me.Lbl_table2.Name = "Lbl_table2"
        Me.Lbl_table2.Size = New System.Drawing.Size(43, 13)
        Me.Lbl_table2.TabIndex = 13
        Me.Lbl_table2.Text = "Table 2"
        '
        'RadioButton1
        '
        Me.RadioButton1.AutoSize = True
        Me.RadioButton1.Checked = True
        Me.RadioButton1.Location = New System.Drawing.Point(102, 9)
        Me.RadioButton1.Name = "RadioButton1"
        Me.RadioButton1.Size = New System.Drawing.Size(36, 17)
        Me.RadioButton1.TabIndex = 14
        Me.RadioButton1.TabStop = True
        Me.RadioButton1.Text = "Or"
        Me.RadioButton1.UseVisualStyleBackColor = True
        '
        'RadioButton2
        '
        Me.RadioButton2.AutoSize = True
        Me.RadioButton2.Location = New System.Drawing.Point(144, 9)
        Me.RadioButton2.Name = "RadioButton2"
        Me.RadioButton2.Size = New System.Drawing.Size(44, 17)
        Me.RadioButton2.TabIndex = 14
        Me.RadioButton2.Text = "And"
        Me.RadioButton2.UseVisualStyleBackColor = True
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Location = New System.Drawing.Point(6, 11)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(90, 13)
        Me.Label5.TabIndex = 15
        Me.Label5.Text = "Boolean Operator"
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Location = New System.Drawing.Point(28, 133)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(102, 13)
        Me.Label6.TabIndex = 15
        Me.Label6.Text = "Comparism Operator"
        '
        'cmbbx_compare
        '
        Me.cmbbx_compare.FormattingEnabled = True
        Me.cmbbx_compare.Items.AddRange(New Object() {"In", "Equals"})
        Me.cmbbx_compare.Location = New System.Drawing.Point(143, 130)
        Me.cmbbx_compare.Name = "cmbbx_compare"
        Me.cmbbx_compare.Size = New System.Drawing.Size(67, 21)
        Me.cmbbx_compare.TabIndex = 16
        Me.ToolTip1.SetToolTip(Me.cmbbx_compare, "Select either the IN or Equality Operator. Note that IN is only suitable for Text" & _
        " Columns")
        '
        'Panel1
        '
        Me.Panel1.Controls.Add(Me.Label5)
        Me.Panel1.Controls.Add(Me.RadioButton1)
        Me.Panel1.Controls.Add(Me.RadioButton2)
        Me.Panel1.Location = New System.Drawing.Point(135, 155)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(200, 44)
        Me.Panel1.TabIndex = 17
        Me.Panel1.Visible = False
        '
        'Form2
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(346, 511)
        Me.Controls.Add(Me.Panel1)
        Me.Controls.Add(Me.cmbbx_compare)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Lbl_table2)
        Me.Controls.Add(Me.lbl_Table1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.chkLbx_columns)
        Me.Controls.Add(Me.btn_remove)
        Me.Controls.Add(Me.btn_merge)
        Me.Controls.Add(Me.btn_Add)
        Me.Controls.Add(Me.cmbbx_columns2)
        Me.Controls.Add(Me.cmbbx_columns1)
        Me.Controls.Add(Me.Lbox_CompareColumns)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow
        Me.Name = "Form2"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Match & Merge"
        Me.Panel1.ResumeLayout(False)
        Me.Panel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Lbox_CompareColumns As System.Windows.Forms.ListBox
    Friend WithEvents cmbbx_columns1 As System.Windows.Forms.ComboBox
    Friend WithEvents cmbbx_columns2 As System.Windows.Forms.ComboBox
    Friend WithEvents btn_Add As System.Windows.Forms.Button
    Friend WithEvents btn_merge As System.Windows.Forms.Button
    Friend WithEvents btn_remove As System.Windows.Forms.Button
    Friend WithEvents chkLbx_columns As System.Windows.Forms.CheckedListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents lbl_Table1 As System.Windows.Forms.Label
    Friend WithEvents Lbl_table2 As System.Windows.Forms.Label
    Friend WithEvents RadioButton1 As System.Windows.Forms.RadioButton
    Friend WithEvents RadioButton2 As System.Windows.Forms.RadioButton
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents cmbbx_compare As System.Windows.Forms.ComboBox
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents ToolTip1 As System.Windows.Forms.ToolTip
End Class
