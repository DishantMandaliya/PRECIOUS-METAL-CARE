Imports System.Data.OleDb
Imports System.Configuration
Public Class stock_details
    Dim a, b, c, d As Double
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim ds As New DataSet
    Public dt1234 As New DataTable
    Public dt As New DataTable
    Private Sub stock_details_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        autogenerate()
        display1()
        dt.Columns.Add("stock_id")
        dt.Columns.Add("category_id")
        dt.Columns.Add("particulars_id")
        dt.Columns.Add("size_id")
        dt.Columns.Add("purity_id")
        dt.Columns.Add("Gross_Weight")
        dt.Columns.Add("Less_Detail")
        dt.Columns.Add("Less_Weight")
        dt.Columns.Add("Net_Weight")
        Try
            load_data.datashow()
            ComboBox4.DataSource = dt10
            ComboBox4.DisplayMember = "category_name"
            ComboBox4.ValueMember = "category_id"

            ComboBox3.DataSource = dt11
            ComboBox3.DisplayMember = "purity_vs"
            ComboBox3.ValueMember = "purity_id"

            ComboBox1.DataSource = dt12
            ComboBox1.DisplayMember = "particular_name"
            ComboBox1.ValueMember = "id"

            ComboBox2.DataSource = dt13
            ComboBox2.DisplayMember = "js_size"
            ComboBox2.ValueMember = "size_id"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        autogenerate()
    End Sub
    Function weight()
        a = Val(TextBox3.Text)
        b = Val(TextBox5.Text)
        d = (a - b)
        TextBox1.Text = d
        Return True
    End Function
    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        weight()
    End Sub
    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        weight()
    End Sub
    Function display()
        Try
            ds.Clear()
            Dim q As String
            q = "select * from stock_table order by stock_id"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(ds)
            DataGridView1.DataSource = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        Return True
    End Function
    Function clear()
        TextBox4.Clear()
        TextBox5.Clear()
        TextBox1.Clear()
        TextBox3.Clear()
        Return True
    End Function
    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (stock_id) from stock_table"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt17)
            Dim data As Integer = dt17.Rows(0).Item(0)
            Label1.Text = data + 1
        Catch ex As Exception
            Label1.Text = "1"
        End Try
        Return True
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click 
        Try
            cmd = New OleDbCommand("insert into stock_table values(@stock_id, @category_id,@particulars_id,@size,@purity,@Gross_Weight,@Less_Detail,@Less_Weight,@Net_Weight)", comm)
            dt.Rows.Add(Label1.Text, ComboBox4.SelectedValue, ComboBox1.SelectedValue, ComboBox2.SelectedValue, ComboBox3.SelectedValue, TextBox3.Text, TextBox4.Text, TextBox5.Text, TextBox1.Text)
            cmd.Parameters.Add("stock_ID", OleDbType.Integer, 10, "stock_ID")
            cmd.Parameters.Add("category_id", OleDbType.Integer, 10, "category_id")
            cmd.Parameters.Add("particulars_id", OleDbType.Integer, 10, "particulars_id")
            cmd.Parameters.Add("size_id", OleDbType.Integer, 10, "size_id")
            cmd.Parameters.Add("purity_id", OleDbType.Integer, 10, "purity_id")
            cmd.Parameters.Add("Gross_Weight", OleDbType.Double, 150, "Gross_Weight")
            cmd.Parameters.Add("Less_Detail", OleDbType.VarChar, 100, "Less_Detail")
            cmd.Parameters.Add("Less_Weight", OleDbType.Double, 100, "Less_Weight")
            cmd.Parameters.Add("Net_Weight", OleDbType.Double, 50, "Net_Weight")
            adp.InsertCommand = cmd
            adp.Update(dt)
            MsgBox("inserted Stock succesfully")
            display1()
            clear()
            autogenerate()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim msgresult As MsgBoxResult = MsgBox("Do you Really want to delete ?", MsgBoxStyle.YesNo)
        If msgresult = MsgBoxResult.Yes Then
            Try
                Dim cmd As New OleDbCommand("delete * from stock_table where stock_id=@stock_id", comm)
                Dim dr As DataRow
                dr = dt1234.Rows(DataGridView1.CurrentRow.Index)
                dr.Delete()
                cmd.Parameters.Add(New OleDbParameter("stock_id", OleDbType.Char, 50, "stock_id"))
                adp.DeleteCommand = cmd
                adp.Update(dt1234)
                MsgBox("Record deleted succesfully")
                clear()
                autogenerate()
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox(ex.Source)
            End Try
        Else : msgresult = MsgBoxResult.No
            MsgBox("Thank You Visit Again!")
        End If
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Label1.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(0)
            ComboBox4.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(1)
            ComboBox1.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(2)
            ComboBox2.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(3)
            ComboBox3.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(4)
            TextBox3.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(5)
            TextBox4.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(6)
            TextBox5.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(7)
            TextBox1.Text = dt1234.Rows(DataGridView1.CurrentRow.Index).Item(8)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub checkconnectionstste()
        If (comm.State = ConnectionState.Open) Then
            comm.Close()
        End If
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            Dim q As String
            checkconnectionstste()
            q = "UPDATE stock_table SET [category_id] = " & Val(ComboBox4.SelectedValue) & ",[particulars_id] = " & Val(ComboBox1.SelectedValue) & ", [size_id] = " & Val(ComboBox2.SelectedValue) & ", [purity_id] = " & Val(ComboBox3.SelectedValue) & ", [Gross_Weight] = " & TextBox3.Text & ", [Less_Detail] = '" & TextBox4.Text & "', [Less_Weight] = " & TextBox5.Text & ", [Net_Weight] = " & TextBox1.Text & " WHERE [stock_id] = " & Val(Label1.Text) & ""
            cmd = New OleDbCommand(q, comm)
            cmd.Connection = comm
            comm.Open()
            cmd.ExecuteNonQuery()
            comm.Close()
            MsgBox("Updated Successfully!")
            display1()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Function display1()
        Try
            dt1234.Clear()
            Dim q As String
            q = "SELECT stock_table.stock_id, category.category_name, particulars.particular_name, Js_size.js_size, purity.purity_vs, stock_table.Gross_Weight, stock_table.Less_Detail, stock_table.Less_Weight, stock_table.Net_Weight FROM purity INNER JOIN (Js_size INNER JOIN (particulars INNER JOIN (category INNER JOIN stock_table ON category.category_id = stock_table.category_id) ON particulars.ID = stock_table.particulars_id) ON Js_size.size_id = stock_table.size_id) ON purity.Purity_ID = stock_table.purity_id order by stock_id"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt1234)
            DataGridView1.DataSource = dt1234
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        Return True
    End Function
    Private Sub Button4_Click_1(sender As Object, e As EventArgs)
        display1()
    End Sub
End Class