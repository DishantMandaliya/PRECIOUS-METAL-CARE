Imports System.Data.OleDb
Imports System.Configuration
Public Class payment
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim dt As New DataTable
    Dim ds As New DataSet
    Dim x As New AutoCompleteStringCollection
    Dim dt123, zaq1, zaq2, dt1 As New DataTable
    Dim q As String
    Dim data As Integer
    Dim a, b, c, d, e1 As Double
    Private Sub payment_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        csname()
        RadioButton2.Checked = True
    End Sub
    Function calculate()
        Try
            a = Val(TextBox1.Text)
            b = Val(TextBox2.Text)
            c = b / a
            TextBox3.Text = c
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return True
    End Function
    Function perc()
        d = 0
        For i = 0 To zaq1.Rows.Count - 1
            a = zaq1.Rows(i).Item(1)
            b = (zaq1.Rows(i).Item(2) / 100)
            c = a * b
            d = d + c
        Next
        Return True
    End Function
    Function pay1()
        e1 = 0
        For i = 0 To zaq2.Rows.Count - 1
            a = Val(zaq2.Rows(i).Item(1))
            e1 = e1 + a
        Next
        Return True
    End Function
    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (payment_id) from payment"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt17)
            data = dt17.Rows(0).Item(0) + 1
        Catch
            data = 1
        End Try
        Return data
    End Function
    Function display()
        zaq1.Clear()
        zaq2.Clear()
        q = "SELECT purchase_master.Supplier_id, purchase_item.Total_weight, purchase_item.purchase_percentage FROM purchase_master INNER JOIN purchase_item ON purchase_master.purchase_masterid = purchase_item.purchase_ID where supplier_id = " & ComboBox1.SelectedIndex + 1 & ";"
        adp = New OleDbDataAdapter(q, comm)
        adp.Fill(zaq1)
        q = "SELECT payment.supplier_id, payment.fine_payment FROM payment where supplier_id = " & ComboBox1.SelectedIndex + 1 & ";"
        Dim adp1 As New OleDbDataAdapter(q, comm)
        adp1.Fill(zaq2)
        Return True
    End Function
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        calculate()
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        autogenerate()
        dt1.Columns.Add("payment_id")
        dt1.Columns.Add("Supplier_id")
        dt1.Columns.Add("Payment_date")
        dt1.Columns.Add("fine_payment")
        Try
            cmd = New OleDbCommand("insert into payment values(@payment_id,@Supplier_id, @Payment_date, @fine_payment)", comm)
            dt1.Rows.Add(data, ComboBox1.SelectedValue, DateTimePicker1.Value, TextBox3.Text)
            cmd.Parameters.Add("payment_id", OleDbType.Integer, 10, "payment_id")
            cmd.Parameters.Add("Supplier_id", OleDbType.Integer, 10, "Supplier_id")
            cmd.Parameters.Add("Payment_date", OleDbType.Date, 10, "Payment_date")
            cmd.Parameters.Add("fine_payment", OleDbType.Double, 20, "fine_payment")
            adp.InsertCommand = cmd
            adp.Update(dt1)
            MsgBox("Payment Successful")
            Label5.Text = 0
            display()
            perc()
            pay1()
            Label5.Text = Val(d) - Val(e1)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        calculate()
    End Sub
    Function csname()
        Try
            q = "select supplier_id, supplier_Name, contact_no from supplier"
            Dim adp1 As New OleDbDataAdapter(q, comm)
            adp1.Fill(dt123)
            ComboBox1.DataSource = dt123
            ComboBox1.DisplayMember = "supplier_Name"
            ComboBox1.ValueMember = "supplier_id"
            For i = 0 To dt123.Rows.Count - 1
                x.Add(dt123.Rows(i)("supplier_Name").ToString)
            Next
            ComboBox1.AutoCompleteMode = AutoCompleteMode.Suggest
            ComboBox1.AutoCompleteSource = AutoCompleteSource.CustomSource
            ComboBox1.AutoCompleteCustomSource = x
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
        Return True
    End Function
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Try
            Label5.Text = 0
            display()
            perc()
            pay1()
            Label5.Text = Val(d) - Val(e1)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = True
    End Sub
    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        TextBox3.Enabled = False
        TextBox1.Enabled = True
        TextBox2.Enabled = True
    End Sub
End Class