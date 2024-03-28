Imports System.Data.OleDb
Imports System.Configuration
Public Class order_Book
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim dt, dt1, dt2 As New DataTable
    Dim ds As New DataSet
    Dim x As New AutoCompleteStringCollection
    Dim dt123 As New DataTable
    Dim q As String
    Private Sub checkconnectionstste()
        If (comm.State = ConnectionState.Open) Then
            comm.Close()
        End If
    End Sub
    Function csname()
        Try
            q = "select customer_id,customer_name,phone_number from customer"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt123)
            ComboBox6.DataSource = dt123
            ComboBox6.DisplayMember = "customer_name"
            ComboBox6.ValueMember = "customer_id"
            For i = 0 To dt123.Rows.Count - 1
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

    Private Sub order_Book_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button4.Enabled = False
        cart()
        insertinto()
        autogenerate()
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
            autogenerate()
            DateTimePicker1.Value = Now
            DateTimePicker2.Value = Now
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function insertinto()
        dt1.Columns.Add("Order_Id")
        dt1.Columns.Add("Order_receive")
        dt1.Columns.Add("Customer_ID")
        dt1.Columns.Add("Description_ds")
        dt2.Columns.Add("Order_detail_id")
        dt2.Columns.Add("Order_due")
        dt2.Columns.Add("category_id")
        dt2.Columns.Add("particulars_id")
        dt2.Columns.Add("size_js")
        dt2.Columns.Add("purity_id")
        dt2.Columns.Add("weight_gms")
        dt2.Columns.Add("Quantity_qt")
        Return True
    End Function
    Function cart()
        Try
            dt.Columns.Add("Order_Id")
            dt.Columns.Add("Order Due")
            dt.Columns.Add("catagory")
            dt.Columns.Add("particulars")
            dt.Columns.Add("size")
            dt.Columns.Add("purity")
            dt.Columns.Add("weight")
            dt.Columns.Add("quantity")
            dt.Columns.Add("Description")
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return True
    End Function
    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (order_id) from order_book"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Update(dt17)
            adp.Fill(dt17)
            Dim data As String = dt17.Rows(0).Item(0)
            Dim d As String = data.Substring(1)
            Dim d1 As String = "O" + Str(d + 1)
            Label1.Text = d1.Replace(" ", "")
        Catch ex As Exception
            Label1.Text = "0101"
        End Try
        Return True
    End Function
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        customer.MdiParent = Homepage
        customer.Show()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim dr, dr1 As DataRow
        dr = dt.Rows(DataGridView2.CurrentRow.Index)
        dr1 = dt2.Rows(DataGridView2.CurrentRow.Index)
        dr1.Delete()
        dr.Delete()
        clear()
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Button4.Enabled = True
        Try
            dt.Rows.Add(Label1.Text, DateTimePicker2.Value, ComboBox1.Text, ComboBox4.Text, ComboBox2.Text, ComboBox3.Text, TextBox3.Text, TextBox4.Text, RichTextBox1.Text)
            dt2.Rows.Add(Label1.Text, DateTimePicker2.Value, ComboBox1.SelectedValue, ComboBox4.SelectedValue, ComboBox2.SelectedValue, ComboBox3.SelectedValue, TextBox3.Text, TextBox4.Text)
            DataGridView2.DataSource = dt
            clear()
            DateTimePicker1.Enabled = False
            ComboBox6.Enabled = False
            TextBox6.Enabled = False
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Function clear()
        TextBox3.Clear()
        TextBox4.Clear()
        ComboBox1.ResetText()
        ComboBox2.ResetText()
        ComboBox3.ResetText()
        ComboBox4.ResetText()
        Return True
    End Function
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try
            Label1.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(0)
            DateTimePicker2.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(1)
            ComboBox1.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(2)
            ComboBox4.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(3)
            ComboBox2.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(4)
            ComboBox3.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(5)
            TextBox3.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(6)
            TextBox4.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(7)
            RichTextBox1.Text = dt.Rows(DataGridView2.CurrentRow.Index).Item(8)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        TextBox6.Text = dt123.Rows(ComboBox6.SelectedIndex).Item(2)
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            cmd = New OleDbCommand("insert into order_book values(@order_id, @order_receive, @customer_id, @Description_ds)", comm)
            dt1.Rows.Add(Label1.Text, DateTimePicker1.Value, ComboBox6.SelectedValue, RichTextBox1.Text)
            cmd.Parameters.Add("order_id", OleDbType.Char, 10, "order_id")
            cmd.Parameters.Add("order_receive", OleDbType.Date, 10, "order_receive")
            cmd.Parameters.Add("customer_id", OleDbType.Integer, 10, "customer_id")
            cmd.Parameters.Add("Description_ds", OleDbType.WChar, 100, "Description_ds")
            adp.InsertCommand = cmd
            adp.Update(dt1)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        For i = 0 To dt.Rows.Count - 1
            Try
                cmd = New OleDbCommand("insert into order_book_details values(@Order_detail_id, @order_due, @category_id, @particulars_id, @size_js, @purity_id, @weight_gms, @Quantity_qt)", comm)
                cmd.Parameters.Add("Order_detail_id", OleDbType.Char, 10, "Order_detail_id")
                cmd.Parameters.Add("order_due", OleDbType.Date, 10, "order_due")
                cmd.Parameters.Add("category_id", OleDbType.Integer, 10, "category_id")
                cmd.Parameters.Add("particulars_id", OleDbType.Integer, 10, "particulars_id")
                cmd.Parameters.Add("size_js", OleDbType.Integer, 10, "size_js")
                cmd.Parameters.Add("purity_id", OleDbType.Integer, 10, "purity_id")
                cmd.Parameters.Add("weight_gms", OleDbType.Double, 10, "weight_gms")
                cmd.Parameters.Add("Quantity_qt", OleDbType.Double, 10, "Quantity_qt")
                adp.InsertCommand = cmd
                adp.Update(dt2)
                MsgBox("Inserted Successfully!")
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox(ex.Source)
            End Try
        Next
        dt.Clear()
        dt1.Clear()
        dt2.Clear()
        autogenerate()
        DateTimePicker1.Enabled = True
        ComboBox6.Enabled = True
        TextBox6.Enabled = True
    End Sub
    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click
        MsgBox(ComboBox1.DisplayMember)
    End Sub

    Private Sub Panel1_Paint(sender As Object, e As PaintEventArgs) Handles Panel1.Paint

    End Sub
End Class