Imports System.Data.OleDb
Imports System.Configuration
Public Class sales
    Dim a, b As Double
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim zaq1, dt1, dt2, dt3, datat As New DataTable
    Public gb As New DataTable
    Dim ds As New DataSet
    Dim x, y As New AutoCompleteStringCollection
    Dim c As Double
    Dim q As String
    Dim i As Integer
    Dim dt123 As New DataTable

    Function csname()
        Try

            q = "select customer_id,customer_name,phone_number from customer"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt123)
            ComboBox6.DataSource = dt123
            ComboBox6.DisplayMember = "customer_name"
            ComboBox6.ValueMember = "customer_id"
            For Me.i = 0 To dt123.Rows.Count - 1
                x.Add(dt123.Rows(i)("customer_name").ToString)
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

    Private Sub sales_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        display1()
        autogenerate()
        stocksummary()
        Button4.Enabled = False


        zaq1.Columns.Add("Category")
        zaq1.Columns.Add("Particulars")
        zaq1.Columns.Add("size")
        zaq1.Columns.Add("purity")
        zaq1.Columns.Add("quantity")
        zaq1.Columns.Add("weight")
        zaq1.Columns.Add("Price per Gram")
        zaq1.Columns.Add("total amount")

        gb.Columns.Add("Category")
        gb.Columns.Add("Particulars")
        'gb.Columns.Add("size")
        'gb.Columns.Add("purity")
        gb.Columns.Add("quantity")
        gb.Columns.Add("Price")
        'gb.Columns.Add("Price per Gram")
        gb.Columns.Add("total amount")

        dt1.Columns.Add("sales_id")
        dt1.Columns.Add("sales_date")
        dt1.Columns.Add("customer_id")
        dt1.Columns.Add("Description_ds")

        dt2.Columns.Add("sales_id")
        dt2.Columns.Add("category_id")
        dt2.Columns.Add("particulars_id")
        dt2.Columns.Add("Size_js")
        dt2.Columns.Add("Purity_id")
        dt2.Columns.Add("Weight_gms")
        dt2.Columns.Add("Quantity_qs")
        dt2.Columns.Add("price_per_gram")
        Try
            load_data.datashow()
            ComboBox1.DataSource = dt10
            ComboBox1.DisplayMember = "category_name"
            ComboBox1.ValueMember = "category_id"
            ComboBox3.DataSource = dt11
            ComboBox3.DisplayMember = "purity_vs"
            ComboBox3.ValueMember = "purity_id"
            ComboBox4.DataSource = dt12
            ComboBox4.DisplayMember = "particular_name"
            ComboBox4.ValueMember = "id"
            ComboBox2.DataSource = dt13
            ComboBox2.DisplayMember = "js_size"
            ComboBox2.ValueMember = "size_id"
            csname()
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function clear()
        TextBox3.Clear()
        TextBox4.Clear()
        TextBox4.Enabled = True
        TextBox1.Clear()
        ComboBox1.ResetText()
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        DateTimePicker1.ResetText()
        Return True
    End Function

    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (sales_id) from sales_item"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt17)
            Dim data As String = dt17.Rows(0).Item(0)
            Dim d As String = data.Substring(1)
            Dim d1 As String = "S" + Str(d + 1)
            Label1.Text = d1.Replace(" ", "")
        Catch
            Label1.Text = "S101"
        End Try
        Return True
    End Function

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        TextBox6.Text = dt123.Rows(ComboBox6.SelectedIndex).Item(2)
    End Sub
    Function add()
        Try
            a = Val(ComboBox3.SelectedItem(2))
            b = Val(TextBox3.Text)
            c = a * b * Val(TextBox1.Text) * Val(TextBox4.Text)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        Return 0
    End Function

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If TextBox1.Text = "" Then
            MsgBox("Please Enter Price Per Gram")
            TextBox1.Focus()
        Else
            Button4.Enabled = True
            ComboBox6.Enabled = False
            TextBox6.Enabled = False
            add()
            gb.Rows.Add(ComboBox3.Text + "K " + ComboBox1.Text, ComboBox4.Text + " (" + TextBox3.Text + " gms) ", TextBox4.Text, TextBox1.Text + " (24k)", c)
            zaq1.Rows.Add(ComboBox1.Text, ComboBox4.Text, ComboBox2.Text, ComboBox3.Text, TextBox4.Text, TextBox3.Text, TextBox1.Text, c)
            dt2.Rows.Add(Label1.Text, ComboBox1.SelectedValue, ComboBox4.SelectedValue, ComboBox2.SelectedValue, ComboBox3.SelectedValue, TextBox4.Text, TextBox3.Text, TextBox1.Text)
            DataGridView1.DataSource = zaq1
            Label18.Text = 0
            For i = 0 To zaq1.Rows.Count - 1
                Label18.Text = Val(Label18.Text) + Val(zaq1.Rows(i).Item(7))
            Next
            clear()
        End If

    End Sub
    Private Sub Button4_Click_1(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            cmd = New OleDbCommand("insert into sales_master values(@sales_id, @sales_date, @customer_id,@description_ds)", comm)
            dt1.Rows.Add(Label1.Text, DateTimePicker1.Value, ComboBox6.SelectedValue, RichTextBox1.Text)
            cmd.Parameters.Add("sales_id", OleDbType.WChar, 10, "sales_id")
            cmd.Parameters.Add("sales_date", OleDbType.Date, 10, "sales_date")
            cmd.Parameters.Add("customer_id", OleDbType.Integer, 10, "customer_id")
            cmd.Parameters.Add("description_ds", OleDbType.WChar, 100, "description_ds")
            adp.InsertCommand = cmd
            adp.Update(dt1)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        For i = 0 To dt1.Rows.Count - 1
            Try
                cmd = New OleDbCommand("insert into sales_item values(@sales_id,@category_id,@particulars_id, @Size_js, @purity_id, @Weight_gms, @Quantity_qs, @price_per_gram)", comm)
                'dt2.Rows.Add(dt2.Rows(i).Item(0), dt2.Rows(i).Item(1), dt2.Rows(i).Item(2), dt2.Rows(i).Item(3), dt2.Rows(i).Item(4), dt2.Rows(i).Item(5), dt2.Rows(i).Item(6))
                cmd.Parameters.Add("sales_id", OleDbType.WChar, 100, "sales_id")
                cmd.Parameters.Add("category_id", OleDbType.Integer, 10, "category_id")
                cmd.Parameters.Add("particulars_id", OleDbType.Integer, 10, "particulars_id")
                cmd.Parameters.Add("Size_js", OleDbType.Integer, 10, "Size_js")
                cmd.Parameters.Add("purity_id", OleDbType.Integer, 10, "purity_id")
                cmd.Parameters.Add("Weight_gms", OleDbType.Double, 10, "Weight_gms")
                cmd.Parameters.Add("Quantity_qs", OleDbType.Double, 10, "Quantity_qs")
                cmd.Parameters.Add("price_per_gram", OleDbType.Double, 10, "price_per_gram")
                adp.InsertCommand = cmd
                adp.Update(dt2)
            Catch ex As Exception
                MsgBox(ex.Message)
            End Try
        Next
        Try
            cmd = New OleDbCommand("insert into Additionalcharges values(@Add_ID, @sales_date1, @particulars_ps1, @quantity_qs1, @price_ps1, @Total_price1)", comm)
            cmd.Parameters.Add("Add_ID", OleDbType.Char, 10, "Add_ID")
            cmd.Parameters.Add("sales_date1", OleDbType.Char, 10, "sales_date1")
            cmd.Parameters.Add("particulars_ps1", OleDbType.WChar, 10, "particulars_ps1")
            cmd.Parameters.Add("quantity_qs1", OleDbType.Integer, 10, "quantity_qs1")
            cmd.Parameters.Add("price_ps1", OleDbType.Double, 20, "price_ps1")
            cmd.Parameters.Add("Total_price1", OleDbType.Double, 20, "Total_price1")
            adp.InsertCommand = cmd
            adp.Update(Additionalcharges.zaq1wsx)
            MsgBox("Inserted Successfully")
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try

        If ComboBox7.SelectedValue <= 1 Then
            MsgBox("Sales Made Other than Stock")
        Else
            Try
                Dim q As String
                checkconnectionstste()
                q = "DELETE * FROM stock_table WHERE [stock_id] = " & ComboBox7.SelectedValue & ""
                cmd = New OleDbCommand(q, comm)
                cmd.Connection = comm
                comm.Open()
                cmd.ExecuteNonQuery()
                comm.Close()
                MsgBox("Deleted Successfully!")
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox(ex.Source)
            End Try
        End If
        autogenerate()
        ds.Clear()
        clear()
        dt1.Clear()
        dt2.Clear()
        zaq1.Clear()
        Additionalcharges.zaq1wsx.Clear()
        autogenerate()
        DateTimePicker1.Enabled = True
        ComboBox6.Enabled = True
        TextBox6.Enabled = True
        RichTextBox1.Clear()
        Label1.Update()
        Label18.Text = "0.00"
        Billings.Show()
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        customer.MdiParent = Homepage
        customer.Show()
    End Sub
    Private Sub checkconnectionstste()
        If (comm.State = ConnectionState.Open) Then
            comm.Close()
        End If
    End Sub

    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Additionalcharges.MdiParent = Homepage
        Additionalcharges.Show()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim dr As DataRow
        dr = zaq1.Rows(DataGridView1.CurrentRow.Index)
        dr.Delete()
    End Sub
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            TextBox4.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(4)
            TextBox1.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(6)
            ComboBox1.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(0)
            ComboBox4.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(1)
            ComboBox2.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(2)
            TextBox3.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(5)
            ComboBox3.Text = zaq1.Rows(DataGridView1.CurrentRow.Index).Item(3)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Function display1()

        Try
            dt3.Clear()
            Dim q As String
            'q = "select * from stock_table"
            q = "SELECT stock_table.stock_id, category.category_name, particulars.particular_name, purity.purity_vs, Js_size.js_size, stock_table.Gross_Weight, stock_table.Less_Detail, stock_table.Less_Weight, stock_table.Net_Weight FROM Js_size INNER JOIN (purity INNER JOIN (particulars INNER JOIN (category INNER JOIN stock_table ON category.category_id = stock_table.category_id) ON particulars.ID = stock_table.particulars_id) ON purity.Purity_ID = stock_table.purity_id) ON Js_size.size_id = stock_table.size_id order by stock_id;"

            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt3)
            ComboBox7.DataSource = dt3
            ComboBox7.DisplayMember = "stock_id"
            ComboBox7.ValueMember = "stock_id"

            For Me.i = 0 To dt3.Rows.Count - 1
                y.Add(dt3.Rows(i)("stock_id").ToString)
            Next
            ComboBox7.AutoCompleteMode = AutoCompleteMode.Suggest
            ComboBox7.AutoCompleteSource = AutoCompleteSource.CustomSource
            ComboBox7.AutoCompleteCustomSource = y
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)

        End Try

        Return True
    End Function
    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        Try
            ComboBox1.Text = dt3.Rows(ComboBox7.SelectedIndex).Item(1)
            ComboBox4.Text = dt3.Rows(ComboBox7.SelectedIndex).Item(2)
            ComboBox2.Text = dt3.Rows(ComboBox7.SelectedIndex).Item(3)
            TextBox3.Text = dt3.Rows(ComboBox7.SelectedIndex).Item(8)
            ComboBox3.Text = dt3.Rows(ComboBox7.SelectedIndex).Item(4)
            TextBox4.Enabled = False
            TextBox4.Text = "1"

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub

    Function stocksummary()
        Try

            sea("SELECT category.category_name, particulars.particular_name, Count(*) AS stock, Sum(stock_table.Net_Weight) AS SumOfNet_Weight FROM category INNER JOIN (particulars INNER JOIN stock_table ON particulars.ID = stock_table.particulars_id) ON category.category_id = stock_table.category_id GROUP BY category.category_name, particulars.particular_name;", datat)
            DataGridView2.DataSource = datat

        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
        Return True
    End Function

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class