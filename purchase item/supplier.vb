Imports System.Data.OleDb
Imports System.Configuration
Public Class supplier
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim dt1 As New DataTable
    Dim ds As New DataSet
    Dim x As New AutoCompleteStringCollection


    Function display()
        Try
            dt1.Clear()
            Dim q As String
            q = "select * from supplier"
            'q = "SELECT supplier.supplier_id, supplier.supplier_Name, supplier.contact_no, supplier.supplier_Address, supplier.Supplier_city, state_select.State_Name FROM state_select INNER JOIN supplier ON state_select.state_ID = supplier.state_id;"
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt1)
            DataGridView1.DataSource = dt1
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return True
    End Function
    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (supplier_id) from supplier "
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt17)
            Dim data As String = dt17.Rows(0).Item(0)
            TextBox12.Text = data + 1
        Catch ex As Exception
            TextBox12.Text = "1"
        End Try
        Return 0
    End Function
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim msgresult As MsgBoxResult = MsgBox("Do you Really want to delete ?", MsgBoxStyle.YesNo)
        If msgresult = MsgBoxResult.Yes Then
            Try
                Dim cmd As New OleDbCommand("delete * from supplier where supplier_id=@supplier_id", comm)
                Dim dr As DataRow
                dr = dt1.Rows(DataGridView1.CurrentRow.Index)
                dr.Delete()
                cmd.Parameters.Add(New OleDbParameter("supplier_id", OleDbType.Integer, 50, "supplier_id"))
                adp.DeleteCommand = cmd
                adp.Update(dt1)
                MsgBox("record deleted successfully")
                clear()
            Catch ex As Exception
                MsgBox(ex.Message)
                MsgBox(ex.Source)
            End Try
        Else : msgresult = MsgBoxResult.No
            MsgBox("Thank You Visit Again!")

        End If
        autogenerate()
    End Sub
    Function clear()
        TextBox12.Clear()
        TextBox5.Clear()
        TextBox11.Clear()
        TextBox6.Clear()
        TextBox13.Clear()
        ComboBox4.Refresh()
        Return True
    End Function

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            cmd = New OleDbCommand("insert into supplier values(@supplier_id, @supplier_Name,@contact_no,@supplier_Address,@Supplier_city,@state_id)", comm)
            dt1.Rows.Add(TextBox12.Text, TextBox5.Text, TextBox11.Text, TextBox6.Text, TextBox13.Text, ComboBox4.SelectedIndex)
            cmd.Parameters.Add("supplier_id", OleDbType.Integer, 50, "supplier_id")
            cmd.Parameters.Add("supplier_Name", OleDbType.Char, 150, "supplier_Name")
            cmd.Parameters.Add("contact_no", OleDbType.Double, 50, "contact_no")
            cmd.Parameters.Add("supplier_Address", OleDbType.VarChar, 150, "supplier_Address")
            cmd.Parameters.Add("Supplier_city", OleDbType.Char, 50, "Supplier_city")
            cmd.Parameters.Add("state_ID", OleDbType.Integer, 50, "state_id")
            adp.InsertCommand = cmd
            adp.Update(dt1)
            MsgBox("inserted Supplier successfully")
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        autogenerate()
    End Sub

    Private Sub supplier_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label2.Hide()
        Try
            load_data.state()
            ComboBox4.DataSource = dt15
            ComboBox4.DisplayMember = "State_Name"
            ComboBox4.ValueMember = "state_id"
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        display()
        autogenerate()
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
            q = "UPDATE supplier SET [supplier_Name] = '" & TextBox5.Text & "', [contact_no] = " & Val(TextBox11.Text) & ", [supplier_Address] = '" & TextBox6.Text & "', [Supplier_city] = '" & TextBox13.Text & "', [state_id] = " & Val(ComboBox4.SelectedValue) & " WHERE [supplier_id] = " & Val(TextBox12.Text) & ""
            cmd = New OleDbCommand(q, comm)
            cmd.Connection = comm
            comm.Open()
            cmd.ExecuteNonQuery()
            comm.Close()
            MsgBox("Updated Successfully!")
            display()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            TextBox12.Text = dt1.Rows(DataGridView1.CurrentRow.Index).Item(0)
            TextBox5.Text = dt1.Rows(DataGridView1.CurrentRow.Index).Item(1)
            TextBox11.Text = dt1.Rows(DataGridView1.CurrentRow.Index).Item(2)
            TextBox6.Text = dt1.Rows(DataGridView1.CurrentRow.Index).Item(3)
            TextBox13.Text = dt1.Rows(DataGridView1.CurrentRow.Index).Item(4)
            ComboBox4.Text = dt1.Rows(DataGridView1.CurrentRow.Index).Item(5)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)

        End Try
    End Sub

    Private Sub supplier_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        purchase.Show()
    End Sub

    Private Sub Panel9_Paint(sender As Object, e As PaintEventArgs) Handles Panel9.Paint

    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        If TextBox11.TextLength < 10 Or TextBox11.TextLength > 10 Then
            Label2.Text = "incorrect"
            Label2.ForeColor = Color.Red
            TextBox11.Focus()
            Label2.Show()
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
        Else
            Label2.Text = "correct"
            Label2.ForeColor = Color.Green
            Label2.Show()
            Button2.Enabled = True
            Button3.Enabled = True
            Button4.Enabled = True
        End If
    End Sub
End Class