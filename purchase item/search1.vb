Imports System.Data.OleDb
Imports System.Configuration
Public Class search1
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim ds As DataSet
    Dim dt, dt1 As New DataTable
    Dim q As String
    Dim d As Integer = 0
    Dim x, y, z, zz As New AutoCompleteStringCollection
    Function display()
        Try
            dt.Clear()
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt)
            DataGridView1.DataSource = dt
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return 0
    End Function
    Public Function showthing()
        If search.ComboBox1.SelectedIndex = 0 Then
            q = "SELECT purchase_master.purchase_masterid, purchase_master.Purchase_date, supplier.supplier_Name, category.category_name, particulars.particular_name, purity.purity_vs, purchase_item.Quantity_qt, purchase_item.purchase_percentage, purchase_item.Total_weight, purchase_master.Description_ds FROM supplier INNER JOIN (purity INNER JOIN (particulars INNER JOIN (category INNER JOIN (purchase_master INNER JOIN purchase_item ON purchase_master.purchase_masterid = purchase_item.purchase_ID) ON category.category_id = purchase_item.category_id) ON particulars.ID = purchase_item.particulars_id) ON purity.Purity_ID = purchase_item.purity_id) ON supplier.supplier_id = purchase_master.Supplier_id;"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 1 Then
            q = "SELECT sales_master.sales_id, sales_master.sales_date, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, sales_item.Weight_gms, sales_item.Quantity_qs, sales_item.price_per_gram, sales_master.description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN ((sales_master INNER JOIN sales_item ON sales_master.sales_id = sales_item.sales_id) INNER JOIN customer ON sales_master.customer_id = customer.customer_id) ON category.category_id = sales_item.category_id) ON particulars.ID = sales_item.particulars_id) ON Js_size.size_id = sales_item.Size_js) ON purity.Purity_ID = sales_item.Purity_id;"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 3 Then
            q = "SELECT payment.payment_id, supplier.supplier_Name, payment.Payment_date, payment.fine_payment FROM supplier INNER JOIN payment ON supplier.supplier_id = payment.supplier_id;"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 2 Then
            q = "SELECT Order_Book.Order_Id, Order_Book.Order_receive, order_book_details.Order_due, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, order_book_details.weight_gms, order_book_details.Quantity_qt, Order_Book.Description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN (customer INNER JOIN (Order_Book INNER JOIN order_book_details ON Order_Book.Order_Id = order_book_details.Order_detail_id) ON customer.customer_id = Order_Book.Customer_ID) ON category.category_id = order_book_details.category_id) ON particulars.ID = order_book_details.particulars_id) ON Js_size.size_id = order_book_details.size_js) ON purity.Purity_ID = order_book_details.purity_id;"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 4 Then
            q = "select * from Customer"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 5 Then
            q = "SELECT supplier.supplier_id, supplier.supplier_Name, supplier.contact_no, supplier.supplier_Address, supplier.Supplier_city, state_select.State_Name FROM state_select INNER JOIN supplier ON state_select.state_ID = supplier.state_id;"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 6 Then
            q = "SELECT stock_table.stock_id, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, stock_table.Gross_Weight, stock_table.Less_Detail, stock_table.Less_Weight, stock_table.Net_Weight FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN stock_table ON category.category_id = stock_table.category_id) ON particulars.ID = stock_table.particulars_id) ON Js_size.size_id = stock_table.size_id) ON purity.Purity_ID = stock_table.purity_id order by stock_id"
            display()
        Else
            TextBox1.Enabled = False
            TextBox2.Enabled = False
            TextBox4.Enabled = False
            TextBox5.Enabled = False
            TextBox3.Enabled = False
        End If
        Return True
    End Function
    Private Sub checkconnectionstste()
        If (comm.State = ConnectionState.Open) Then
            comm.Close()
        End If
    End Sub
    Private Sub search1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        showthing()
        load_Data.datashow()
    End Sub
    Private Sub search1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        search.Show()
    End Sub
    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        DataGridView1.DataSource = load_data.dt11
        autogenerate(dt11)
        TextBox1.Enabled = False
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        TextBox4.Enabled = False
        TextBox5.Enabled = False
    End Sub
    Private Sub RadioButton4_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton4.CheckedChanged
        DataGridView1.DataSource = load_data.dt12
        autogenerate(dt12)
        TextBox4.Enabled = False
        TextBox1.Enabled = True
        TextBox2.Enabled = False
        TextBox5.Enabled = False
        TextBox3.Enabled = False
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        DataGridView1.DataSource = load_data.dt10
        autogenerate(dt10)
        TextBox1.Enabled = False
        TextBox3.Enabled = False
        TextBox2.Enabled = False
        TextBox4.Enabled = True
        TextBox5.Enabled = False
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        DataGridView1.DataSource = load_data.dt13
        autogenerate(dt13)
        TextBox2.Enabled = False
        TextBox1.Enabled = False
        TextBox4.Enabled = False
        TextBox5.Enabled = True
        TextBox3.Enabled = False
    End Sub
    Function auto()
        For i = 0 To load_Data.dt10.Rows.Count - 1
            x.Add(load_data.dt10.Rows(i)("category_name").ToString)
        Next
        TextBox4.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox4.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox4.AutoCompleteCustomSource = x

        For i = 0 To load_Data.dt11.Rows.Count - 1
            y.Add(load_data.dt11.Rows(i)("purity_vs").ToString)
        Next
        TextBox6.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox6.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox6.AutoCompleteCustomSource = y

        For i = 0 To load_Data.dt12.Rows.Count - 1
            z.Add(load_data.dt12.Rows(i)("particular_name").ToString)
        Next
        TextBox4.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox4.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox4.AutoCompleteCustomSource = z

        For i = 0 To load_Data.dt13.Rows.Count - 1
            zz.Add(load_Data.dt13.Rows(i)("js_size").ToString)
        Next
        TextBox5.AutoCompleteMode = AutoCompleteMode.Suggest
        TextBox5.AutoCompleteSource = AutoCompleteSource.CustomSource
        TextBox5.AutoCompleteCustomSource = zz
        Return True
    End Function


    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text.Count >= 1 Then
            If search.ComboBox1.SelectedIndex = 0 Then
                q = "SELECT purchase_master.purchase_masterid, purchase_master.Purchase_date, supplier.supplier_Name, category.category_name, particulars.particular_name, purity.purity_vs, purchase_item.Quantity_qt, purchase_item.purchase_percentage, purchase_item.Total_weight, purchase_master.Description_ds FROM supplier INNER JOIN (purity INNER JOIN (particulars INNER JOIN (category INNER JOIN (purchase_master INNER JOIN purchase_item ON purchase_master.purchase_masterid = purchase_item.purchase_ID) ON category.category_id = purchase_item.category_id) ON particulars.ID = purchase_item.particulars_id) ON purity.Purity_ID = purchase_item.purity_id) ON supplier.supplier_id = purchase_master.Supplier_id where purchase_masterid = '" + TextBox1.Text + "'"
                display()
            ElseIf search.ComboBox1.SelectedIndex = 1 Then
                q = "SELECT sales_master.sales_id, sales_master.sales_date, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, sales_item.Weight_gms, sales_item.Quantity_qs, sales_item.price_per_gram, sales_master.description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN ((sales_master INNER JOIN sales_item ON sales_master.sales_id = sales_item.sales_id) INNER JOIN customer ON sales_master.customer_id = customer.customer_id) ON category.category_id = sales_item.category_id) ON particulars.ID = sales_item.particulars_id) ON Js_size.size_id = sales_item.Size_js) ON purity.Purity_ID = sales_item.Purity_id where sales_master.sales_id = '" + TextBox1.Text + "'"
                display()
            ElseIf search.ComboBox1.SelectedIndex = 3 Then
                q = "SELECT payment.payment_id, supplier.supplier_Name, payment.Payment_date, payment.fine_payment FROM supplier INNER JOIN payment ON supplier.supplier_id = payment.supplier_id where payment_id = " + TextBox1.Text + ""
                display()
            ElseIf search.ComboBox1.SelectedIndex = 2 Then
                q = "SELECT Order_Book.Order_Id, Order_Book.Order_receive, order_book_details.Order_due, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, order_book_details.weight_gms, order_book_details.Quantity_qt, Order_Book.Description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN (customer INNER JOIN (Order_Book INNER JOIN order_book_details ON Order_Book.Order_Id = order_book_details.Order_detail_id) ON customer.customer_id = Order_Book.Customer_ID) ON category.category_id = order_book_details.category_id) ON particulars.ID = order_book_details.particulars_id) ON Js_size.size_id = order_book_details.size_js) ON purity.Purity_ID = order_book_details.purity_id where Order_Id = '" + TextBox1.Text + "'"
                display()
            ElseIf search.ComboBox1.SelectedIndex = 4 Then
                q = "select * from Customer where customer_id = " + TextBox1.Text + ""
                display()
            ElseIf search.ComboBox1.SelectedIndex = 5 Then
                q = "SELECT supplier.supplier_id, supplier.supplier_Name, supplier.contact_no, supplier.supplier_Address, supplier.Supplier_city, state_select.State_Name FROM state_select INNER JOIN supplier ON state_select.state_ID = supplier.state_id where supplier_id = " + TextBox1.Text + ""
                display()
            ElseIf search.ComboBox1.SelectedIndex = 6 Then
                q = "SELECT stock_table.stock_id, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, stock_table.Gross_Weight, stock_table.Less_Detail, stock_table.Less_Weight, stock_table.Net_Weight FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN stock_table ON category.category_id = stock_table.category_id) ON particulars.ID = stock_table.particulars_id) ON Js_size.size_id = stock_table.size_id) ON purity.Purity_ID = stock_table.purity_id where Stock_ID = " + TextBox1.Text + " order by stock_id"
                display()
            End If
        ElseIf TextBox1.Text.Count = 0 Then
            showthing()
        End If
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If search.ComboBox1.SelectedIndex = 0 Then
            q = "SELECT purchase_master.purchase_masterid, purchase_master.Purchase_date, supplier.supplier_Name, category.category_name, particulars.particular_name, purity.purity_vs, purchase_item.Quantity_qt, purchase_item.purchase_percentage, purchase_item.Total_weight, purchase_master.Description_ds FROM supplier INNER JOIN (purity INNER JOIN (particulars INNER JOIN (category INNER JOIN (purchase_master INNER JOIN purchase_item ON purchase_master.purchase_masterid = purchase_item.purchase_ID) ON category.category_id = purchase_item.category_id) ON particulars.ID = purchase_item.particulars_id) ON purity.Purity_ID = purchase_item.purity_id) ON supplier.supplier_id = purchase_master.Supplier_id where supplier.supplier_Name like '" + TextBox4.Text + "%'"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 1 Then
            q = "SELECT sales_master.sales_id, sales_master.sales_date, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, sales_item.Weight_gms, sales_item.Quantity_qs, sales_item.price_per_gram, sales_master.description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN ((sales_master INNER JOIN sales_item ON sales_master.sales_id = sales_item.sales_id) INNER JOIN customer ON sales_master.customer_id = customer.customer_id) ON category.category_id = sales_item.category_id) ON particulars.ID = sales_item.particulars_id) ON Js_size.size_id = sales_item.Size_js) ON purity.Purity_ID = sales_item.Purity_id where customer.customer_name like '" + TextBox4.Text + "%'"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 3 Then
            q = "SELECT payment.payment_id, supplier.supplier_Name, payment.Payment_date, payment.fine_payment FROM supplier INNER JOIN payment ON supplier.supplier_id = payment.supplier_id where payment_id = '" + TextBox1.Text + "'"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 2 Then
            q = "SELECT Order_Book.Order_Id, Order_Book.Order_receive, order_book_details.Order_due, customer.customer_name, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, order_book_details.weight_gms, order_book_details.Quantity_qt, Order_Book.Description_ds FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN (customer INNER JOIN (Order_Book INNER JOIN order_book_details ON Order_Book.Order_Id = order_book_details.Order_detail_id) ON customer.customer_id = Order_Book.Customer_ID) ON category.category_id = order_book_details.category_id) ON particulars.ID = order_book_details.particulars_id) ON Js_size.size_id = order_book_details.size_js) ON purity.Purity_ID = order_book_details.purity_id where customer.customer_name like '" + TextBox4.Text + "%'"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 4 Then
            q = "select * from Customer where customer_name like '" + TextBox4.Text + "%'"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 5 Then
            q = "SELECT supplier.supplier_id, supplier.supplier_Name, supplier.contact_no, supplier.supplier_Address, supplier.Supplier_city, state_select.State_Name FROM state_select INNER JOIN supplier ON state_select.state_ID = supplier.state_id where supplier.supplier_Name like '" + TextBox4.Text + "%'"
            display()
        ElseIf search.ComboBox1.SelectedIndex = 6 Then
            q = "SELECT stock_table.stock_id, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, stock_table.Gross_Weight, stock_table.Less_Detail, stock_table.Less_Weight, stock_table.Net_Weight FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN stock_table ON category.category_id = stock_table.category_id) ON particulars.ID = stock_table.particulars_id) ON Js_size.size_id = stock_table.size_id) ON purity.Purity_ID = stock_table.purity_id where category.category_name like '" + TextBox4.Text + "%' "
            display()
        End If
    End Sub

   

    Function autogenerate(dt As DataTable)
        For i = 0 To dt.Rows.Count - 1
            d = (Math.Max(d, dt.Rows(i).Item(0)))
        Next
        d = d + 1
        Return 0
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        TextBox1.Clear()
        If Button1.Text = "Add" Then
            If search.ComboBox1.SelectedIndex = 7 Then
                If RadioButton1.Checked = True Then
                    'category
                    checkconnectionstste()
                    Try

                        cmd.Connection = comm
                        cmd.CommandText = "insert into category values (" & d & "," & TextBox5.Text & ") "
                        comm.Open()
                        cmd.ExecuteNonQuery()
                        comm.Close()
                        MsgBox("Inserted Successfully!")
                        DataGridView1.DataSource = dt10
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try

                ElseIf RadioButton2.Checked = True Then
                    'size
                    checkconnectionstste()
                    Try

                        cmd.Connection = comm
                        cmd.CommandText = "insert into js_size values (" & d & "," & TextBox5.Text & ") "
                        comm.Open()
                        cmd.ExecuteNonQuery()
                        comm.Close()
                        MsgBox("Inserted Successfully!")
                        DataGridView1.DataSource = dt13
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                ElseIf RadioButton3.Checked = True Then
                    'purity
                    Try
                        checkconnectionstste()
                        cmd.Connection = comm
                        cmd.CommandText = "insert into purity(purity_vs,purity_value) values (" & d & ",'" + TextBox2.Text + "', '" + TextBox3.Text + "') "
                        comm.Open()
                        cmd.ExecuteNonQuery()
                        comm.Close()
                        MsgBox("Inserted Successfully!")
                        DataGridView1.DataSource = dt11
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                Else
                    'particulars
                    Try
                        checkconnectionstste()
                        cmd.Connection = comm
                        cmd.CommandText = "insert into particulars(particular_name) values (" & d & ",'" + TextBox1.Text + "') "
                        comm.Open()
                        cmd.ExecuteNonQuery()
                        comm.Close()
                        MsgBox("Inserted Successfully!")
                        DataGridView1.DataSource = dt12
                    Catch ex As Exception
                        MsgBox(ex.Message)
                    End Try
                End If
            End If
        ElseIf Button1.Text = "Delete" Then
            If search.ComboBox1.SelectedIndex = 0 Then
                'purchase
                Try
                    Dim cmd As New OleDbCommand("delete * from purchase_master where purchase_masterid=@purchase_masterid", comm)
                    Dim dr As DataRow
                    dr = dt.Rows(DataGridView1.CurrentRow.Index)
                    dr.Delete()
                    cmd.Parameters.Add(New OleDbParameter("purchase_masterid", OleDbType.WChar, 255, "purchase_masterid"))
                    adp.DeleteCommand = cmd
                    adp.Update(dt)
                    MsgBox("record deleted succesfully")
                Catch ex As Exception
                    MsgBox(ex.Message)
                    MsgBox(ex.Source)
                End Try
            ElseIf search.ComboBox1.SelectedIndex = 1 Then
                'sales
                Try
                    Dim cmd As New OleDbCommand("delete * from sales_master where sales_id=@sales_id", comm)
                    Dim dr As DataRow
                    dr = dt.Rows(DataGridView1.CurrentRow.Index)
                    dr.Delete()
                    cmd.Parameters.Add(New OleDbParameter("sales_id", OleDbType.WChar, 255, "sales_id"))
                    adp.DeleteCommand = cmd
                    adp.Update(dt)
                    MsgBox("record deleted succesfully")
                Catch ex As Exception
                    MsgBox(ex.Message)
                    MsgBox(ex.Source)
                End Try
            ElseIf search.ComboBox1.SelectedIndex = 2 Then
                'orderbook
                Try
                    Dim cmd As New OleDbCommand("delete * from Order_Book where Order_Id=@Order_Id", comm)
                    Dim dr As DataRow
                    dr = dt.Rows(DataGridView1.CurrentRow.Index)
                    dr.Delete()
                    cmd.Parameters.Add(New OleDbParameter("Order_Id", OleDbType.WChar, 255, "Order_Id"))
                    adp.DeleteCommand = cmd
                    adp.Update(dt)
                    MsgBox("record deleted succesfully")
                Catch ex As Exception
                    MsgBox(ex.Message)
                    MsgBox(ex.Source)
                End Try
            ElseIf search.ComboBox1.SelectedIndex = 3 Then
                'payment
                Try
                    Dim cmd As New OleDbCommand("delete * from payment where payment_id=@payment_id", comm)
                    Dim dr As DataRow
                    dr = dt.Rows(DataGridView1.CurrentRow.Index)
                    dr.Delete()
                    cmd.Parameters.Add(New OleDbParameter("payment_id", OleDbType.Integer, 255, "payment_id"))
                    adp.DeleteCommand = cmd
                    adp.Update(dt)
                    MsgBox("record deleted succesfully")
                Catch ex As Exception
                    MsgBox(ex.Message)
                    MsgBox(ex.Source)
                End Try

            ElseIf search.ComboBox1.SelectedIndex = 7 Then
                If RadioButton1.Checked = True Then
                    'category
                    Try
                        Dim cmd As New OleDbCommand("delete * from category where category_id=@category_id", comm)
                        Dim dr As DataRow
                        dr = dt10.Rows(DataGridView1.CurrentRow.Index)
                        dr.Delete()
                        cmd.Parameters.Add(New OleDbParameter("category_id", OleDbType.Integer, 10, "category_id"))
                        adp.DeleteCommand = cmd
                        adp.Update(dt10)
                        MsgBox("record deleted succesfully")
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        MsgBox(ex.Source)
                    End Try
                ElseIf RadioButton2.Checked = True Then
                    'size
                    Try
                        Dim cmd As New OleDbCommand("delete * from js_size where size_id=@size_id", comm)
                        Dim dr As DataRow
                        dr = dt13.Rows(DataGridView1.CurrentRow.Index)
                        dr.Delete()
                        cmd.Parameters.Add(New OleDbParameter("size_id", OleDbType.Integer, 10, "size_id"))
                        adp.DeleteCommand = cmd
                        adp.Update(dt13)
                        MsgBox("record deleted succesfully")
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        MsgBox(ex.Source)
                    End Try
                ElseIf RadioButton3.Checked = True Then
                    'purity
                    Try
                        Dim cmd As New OleDbCommand("delete * from Purity where Purity_ID=@Purity_ID", comm)
                        Dim dr As DataRow
                        dr = dt11.Rows(DataGridView1.CurrentRow.Index)
                        dr.Delete()
                        cmd.Parameters.Add(New OleDbParameter("Purity_ID", OleDbType.Integer, 10, "Purity_ID"))
                        adp.DeleteCommand = cmd
                        adp.Update(dt11)
                        MsgBox("record deleted succesfully")
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        MsgBox(ex.Source)
                    End Try
                Else
                    'particulars
                    Try
                        Dim cmd As New OleDbCommand("delete * from particulars where ID=@ID", comm)
                        Dim dr As DataRow
                        dr = dt12.Rows(DataGridView1.CurrentRow.Index)
                        dr.Delete()
                        cmd.Parameters.Add(New OleDbParameter("ID", OleDbType.Integer, 10, "ID"))
                        adp.DeleteCommand = cmd
                        adp.Update(dt12)
                        MsgBox("record deleted succesfully")
                    Catch ex As Exception
                        MsgBox(ex.Message)
                        MsgBox(ex.Source)
                    End Try
                End If
            End If
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            If Button1.Text = "Delete" Then
                If search.ComboBox1.SelectedIndex = 0 Then
                    'purchase
                    TextBox1.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(0)

                ElseIf search.ComboBox1.SelectedIndex = 1 Then
                    'sales
                    TextBox1.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(0)
                ElseIf search.ComboBox1.SelectedIndex = 2 Then
                    'orderbook
                    TextBox1.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(0)
                ElseIf search.ComboBox1.SelectedIndex = 3 Then
                    'payment
                    TextBox1.Text = dt.Rows(DataGridView1.CurrentRow.Index).Item(0)
                
                ElseIf search.ComboBox1.SelectedIndex = 7 Then
                    If RadioButton1.Checked = True Then
                        'category

                        TextBox1.Text = dt10.Rows(DataGridView1.CurrentRow.Index).Item(0)
                    ElseIf RadioButton2.Checked = True Then
                        'size                
                        TextBox1.Text = dt13.Rows(DataGridView1.CurrentRow.Index).Item(0)
                    ElseIf RadioButton3.Checked = True Then
                        'purity
                        TextBox1.Text = dt11.Rows(DataGridView1.CurrentRow.Index).Item(0)
                    Else
                        'particulars
                        TextBox1.Text = dt12.Rows(DataGridView1.CurrentRow.Index).Item(0)
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        If search.ComboBox1.SelectedIndex = 6 Then
            q = "SELECT stock_table.stock_id, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, stock_table.Gross_Weight, stock_table.Less_Detail, stock_table.Less_Weight, stock_table.Net_Weight FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN stock_table ON category.category_id = stock_table.category_id) ON particulars.ID = stock_table.particulars_id) ON Js_size.size_id = stock_table.size_id) ON purity.Purity_ID = stock_table.purity_id where particulars.particular_name like '" + TextBox5.Text + "%' "
            display()
        End If
    End Sub
End Class