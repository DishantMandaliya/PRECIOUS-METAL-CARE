Imports System.Data.OleDb
Imports System.Configuration
Public Class view_order
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim d, dt As New DataTable
    Dim ds As New DataSet
    Dim cmd As New OleDbCommand
    Private Sub view_order_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            load_data.datashow()
            sea("SELECT Order_Book.Order_Id, Order_Book.Order_receive, order_book_details.Order_due, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, order_book_details.weight_gms, order_book_details.Quantity_qt, Order_Book.Description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN (customer INNER JOIN (Order_Book INNER JOIN order_book_details ON Order_Book.Order_Id = order_book_details.Order_detail_id) ON customer.customer_id = Order_Book.Customer_ID) ON category.category_id = order_book_details.category_id) ON particulars.ID = order_book_details.particulars_id) ON Js_size.size_id = order_book_details.size_js) ON purity.Purity_ID = order_book_details.purity_id order by order_book_details.Order_due;", d)
            DataGridView1.DataSource = d
            ComboBox1.DataSource = d
            ComboBox1.DisplayMember = "Order_id"
            ComboBox1.ValueMember = "Order_id"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        sea("SELECT Order_Book.Order_Id, Order_Book.Order_receive, order_book_details.Order_due, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, order_book_details.weight_gms, order_book_details.Quantity_qt, Order_Book.Description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN (customer INNER JOIN (Order_Book INNER JOIN order_book_details ON Order_Book.Order_Id = order_book_details.Order_detail_id) ON customer.customer_id = Order_Book.Customer_ID) ON category.category_id = order_book_details.category_id) ON particulars.ID = order_book_details.particulars_id) ON Js_size.size_id = order_book_details.size_js) ON purity.Purity_ID = order_book_details.purity_id where Order_Id = '" + ComboBox1.Text + "'", dt)
        DataGridView1.DataSource = dt
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        RichTextBox1.Text = d.Rows(DataGridView1.CurrentRow.Index).Item(10)
    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs) Handles TableLayoutPanel1.Paint

    End Sub
End Class