Imports System.Data.OleDb
Imports System.Configuration

Public Class purchase
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim dt, dt1, dt2, dt112, dt1234 As New DataTable
    Dim ds As New DataSet
    Dim x As New AutoCompleteStringCollection
    Dim dt123 As New DataTable
    Dim q As String
    Dim a, b, c As Double
    Function insertinto()
        dt1.Columns.Add("purchase_masterid")
        dt1.Columns.Add("Purchase_date")
        dt1.Columns.Add("Supplier_id")
        dt1.Columns.Add("Description_ds")
        dt2.Columns.Add("purchase_ID")
        dt2.Columns.Add("category_id")
        dt2.Columns.Add("particulars_id")
        dt2.Columns.Add("purity_id")
        dt2.Columns.Add("Quantity_qt")
        dt2.Columns.Add("purchase_percentage")
        dt2.Columns.Add("Total_weight")
        Return True
    End Function
    Function csname()
        Try
            q = "select supplier_id, supplier_Name, contact_no from supplier"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt123)
            ComboBox6.DataSource = dt123
            ComboBox6.DisplayMember = "supplier_Name"
            ComboBox6.ValueMember = "supplier_id"
            For i = 0 To dt123.Rows.Count - 1
                x.Add(dt123.Rows(i)("supplier_Name").ToString)
            Next
            ComboBox6.AutoCompleteMode = AutoCompleteMode.Suggest
            ComboBox6.AutoCompleteSource = AutoCompleteSource.CustomSource
            ComboBox6.AutoCompleteCustomSource = x
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
        Return True
    End Function
    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (purchase_masterid) from purchase_master"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt17)
            Dim data As String = dt17.Rows(0).Item(0)
            Dim d As String = data.Substring(1)
            Dim d1 As String = "P" + Str(d + 1)
            Label1.Text = d1.Replace(" ", "")
        Catch ex As Exception
            Label1.Text = "P101"
        End Try
        Return True
    End Function
    Private Sub purchase_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        insertinto()
        dt.Columns.Add("Category")
        dt.Columns.Add("Particulars")
        dt.Columns.Add("purity")
        dt.Columns.Add("quantity")
        dt.Columns.Add("Purchase Percentage")
        dt.Columns.Add("Gross weight")
        dt.Columns.Add("Net Payable (Fine/Amount)")
        load_data.datashow()
        ComboBox1.DataSource = dt10
        ComboBox1.DisplayMember = "category_name"
        ComboBox1.ValueMember = "category_id"

        ComboBox3.DataSource = dt11
        ComboBox3.DisplayMember = "purity_vs"
        ComboBox3.ValueMember = "purity_id"

        ComboBox4.DataSource = dt12
        ComboBox4.DisplayMember = "particular_name"
        ComboBox4.ValueMember = "ID"
        autogenerate()
        csname()
    End Sub
    Function clear()
        TextBox5.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox6.Clear()
        ComboBox1.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        ComboBox6.ResetText()
        Return True
    End Function
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Me.Hide()
        supplier.MdiParent = Homepage
        supplier.Show()
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs)
        payment.Show()
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        TextBox6.Text = dt123.Rows(ComboBox6.SelectedIndex).Item(2)
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            ComboBox1.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(0)
            ComboBox4.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(1)
            TextBox3.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(5)
            ComboBox3.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(2)
            TextBox4.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(3)
            TextBox5.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(4)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '==========Add To Cart button click===================
        dt.Rows.Add(ComboBox1.Text, ComboBox4.Text, ComboBox3.Text, TextBox4.Text, TextBox5.Text, TextBox3.Text, c)
        dt2.Rows.Add(Label1.Text, ComboBox1.SelectedValue, ComboBox4.SelectedValue, ComboBox3.SelectedValue, TextBox4.Text, TextBox5.Text, TextBox3.Text)
        DataGridView1.DataSource = dt
        Label18.Text = 0
        For i = 0 To dt.Rows.Count - 1
            Label18.Text = Val(Label18.Text) + Val(dt.Rows(i).Item(6))
        Next
        DateTimePicker1.Enabled = False
        ComboBox6.Enabled = False
        TextBox6.Enabled = False
        TextBox3.Clear()
        TextBox4.Clear()
    End Sub
    Function perc()
        a = Val(TextBox5.Text)
        b = (Val(TextBox3.Text) / 100)
        c = a * b
        Return True
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        '=============Remove Button Click ==============
        Dim dr As DataRow
        dr = dt.Rows(DataGridView1.CurrentRow.Index)
        dr.Delete()
        Label18.Text = 0
        For i = 0 To dt.Rows.Count - 1
            Label18.Text = Val(Label18.Text) + Val(dt.Rows(i).Item(9))
        Next
        clear()
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        '==========Finish Button CLick===========
        Try
            cmd = New OleDbCommand("insert into purchase_master values(@purchase_masterid, @Purchase_date, @Supplier_id,@Description_ds)", comm)
            dt1.Rows.Add(Label1.Text, DateTimePicker1.Value, ComboBox6.SelectedValue, RichTextBox1.Text)
            cmd.Parameters.Add("purchase_masterid", OleDbType.Char, 10, "purchase_masterid")
            cmd.Parameters.Add("Purchase_date", OleDbType.Date, 10, "Purchase_date")
            cmd.Parameters.Add("Supplier_id", OleDbType.Integer, 10, "Supplier_id")
            cmd.Parameters.Add("Description_ds", OleDbType.WChar, 100, "Description_ds")
            adp.InsertCommand = cmd
            adp.Update(dt1)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        For i = 0 To dt1.Rows.Count - 1
            Try
                cmd = New OleDbCommand("insert into purchase_item values(@purchase_ID,@category_id,@particulars_id,@purity_id,@Quantity_qt,@purchase_percentage,@Total_weight)", comm)
                cmd.Parameters.Add("purchase_ID", OleDbType.Char, 10, "purchase_ID")
                cmd.Parameters.Add("category_id", OleDbType.Char, 10, "category_id")
                cmd.Parameters.Add("particulars_id", OleDbType.Char, 10, "particulars_id")
                cmd.Parameters.Add("purity_id", OleDbType.Integer, 10, "purity_id")
                cmd.Parameters.Add("Quantity_qt", OleDbType.Double, 10, "Quantity_qt")
                cmd.Parameters.Add("purchase_percentage", OleDbType.Double, 10, "purchase_percentage")
                cmd.Parameters.Add("Total_weight", OleDbType.Double, 10, "Total_weight")
                adp.InsertCommand = cmd
                adp.Update(dt2)
                adp.Update(dt1)
                MsgBox("Inserted Successfully!")
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Next
        ds.Clear()
        clear()
        dt.Clear()
        dt1.Clear()
        dt2.Clear()
        Label18.Text = "0.00"
        autogenerate()
        DateTimePicker1.Enabled = True
        ComboBox6.Enabled = True
        TextBox6.Enabled = True
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        perc()
    End Sub
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        perc()
    End Sub
    Private Sub Button3_Click_1(sender As Object, e As EventArgs) Handles Button3.Click
        payment.Show()
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs)
        Try
            dt.Clear()
            dt112.Clear()
            dt1234.Clear()
            Dim adp As New OleDbDataAdapter("Select * from purchase_master where purchase_masterid = '" + Label1.Text + "' ", comm)
            adp.Fill(dt112)
            Dim adp1 As New OleDbDataAdapter("SELECT category.category_name, particulars.particular_name, purity.purity_vs, purchase_item.Quantity_qt, purchase_item.purchase_percentage, purchase_item.Total_weight FROM purity INNER JOIN (particulars INNER JOIN (category INNER JOIN purchase_item ON category.category_id = purchase_item.category_id) ON particulars.ID = purchase_item.particulars_id) ON purity.Purity_ID = purchase_item.purity_id where purchase_ID = '" + Label1.Text + "' ", comm)
            adp1.Fill(dt1234)
            For i = 0 To dt1234.Rows.Count - 1
                dt.Rows.Add(dt1234.Rows(i).Item(0), dt1234.Rows(i).Item(1), dt1234.Rows(i).Item(2), dt1234.Rows(i).Item(3), dt1234.Rows(i).Item(4), dt1234.Rows(i).Item(5))
            Next
            DataGridView1.DataSource = dt
            Label1.Enabled = False
            DateTimePicker1.Enabled = False
            ComboBox6.SelectedValue = dt112.Rows(0).Item(2)
            DateTimePicker1.Value = dt112.Rows(0).Item(1)
            TextBox6.Enabled = False
            RichTextBox1.Text = dt112.Rows(0).Item(3)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
End Class
