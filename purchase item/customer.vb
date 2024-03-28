Imports System.Data.OleDb
Imports System.Configuration
Public Class customer
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim dt As New DataTable
    Dim ds As New DataSet
    Dim cmd As New OleDbCommand
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            cmd = New OleDbCommand("insert into customer values(@customer_id,@customer_name,@Address,@phone_Number,@references_rf)", comm)
            ds.Tables(0).Rows.Add(TextBox12.Text, TextBox5.Text, TextBox11.Text, TextBox6.Text, TextBox13.Text)
            cmd.Parameters.Add("customer_id", OleDbType.Integer, 50, "customer_id")
            cmd.Parameters.Add("customer_name", OleDbType.Char, 150, "customer_name")
            cmd.Parameters.Add("Address_ad", OleDbType.VarChar, 250, "Address_ad")
            cmd.Parameters.Add("phone_Number", OleDbType.Double, 100, "phone_Number")
            cmd.Parameters.Add("references_rf", OleDbType.Char, 50, "references_rf")
            adp.InsertCommand = cmd
            adp.Update(ds)
            MsgBox("inserted customer succesfully")
            clear()
            autogenerate()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Function display()
        Try
            ds.Clear()
            Dim q As String
            q = "Select * From customer order by customer_id"
            adp = New OleDbDataAdapter(q, comm)
            adp.Fill(ds)
            DataGridView2.DataSource = ds.Tables(0)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        Return True
    End Function
    Private Sub customer_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        
        Label1.Hide()
        display()
        autogenerate()
    End Sub
    Function autogenerate()
        Try
            Dim dt17 As New DataTable
            Dim q As String
            q = "select max (customer_ID) from customer "
            Dim adp As New OleDbDataAdapter(q, comm)
            adp.Fill(dt17)
            Dim data As String = dt17.Rows(0).Item(0)
            TextBox12.Text = data + 1
        Catch ex As Exception
            TextBox12.Text = "1"
        End Try
        Return 0
    End Function
    Private Sub checkconnectionstste()
        If (comm.State = ConnectionState.Open) Then
            comm.Close()
        End If
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            Dim q As String
            checkconnectionstste()
            q = "UPDATE customer SET [customer_name] = '" & TextBox5.Text & "', [Address_ad] = '" & TextBox11.Text & "', [PHONE_NUMBER] = " & Val(TextBox6.Text) & ", [references_rf] = '" & TextBox13.Text & "' WHERE [customer_id] = " & Val(TextBox12.Text) & ""
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
    Private Sub DataGridView2_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView2.CellClick
        Try
            TextBox12.Text = ds.Tables(0).Rows(DataGridView2.CurrentRow.Index).Item(0)
            TextBox5.Text = ds.Tables(0).Rows(DataGridView2.CurrentRow.Index).Item(1)
            TextBox11.Text = ds.Tables(0).Rows(DataGridView2.CurrentRow.Index).Item(2)
            TextBox6.Text = ds.Tables(0).Rows(DataGridView2.CurrentRow.Index).Item(3)
            TextBox13.Text = ds.Tables(0).Rows(DataGridView2.CurrentRow.Index).Item(4)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim msgresult As MsgBoxResult = MsgBox("Do you Really want to delete ?", MsgBoxStyle.YesNo)
        If msgresult = MsgBoxResult.Yes Then
            Try
                Dim cmd As New OleDbCommand("delete * from customer where customer_id=@customer_id", comm)
                Dim dr As DataRow
                dr = ds.Tables(0).Rows(DataGridView2.CurrentRow.Index)
                dr.Delete()
                cmd.Parameters.Add(New OleDbParameter("customer_id", OleDbType.Integer, 50, "customer_id"))
                adp.DeleteCommand = cmd
                adp.Update(ds.Tables(0))
                MsgBox("record deleted succesfully")
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
        Return True
    End Function
    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        If TextBox6.TextLength < 10 Or TextBox6.TextLength > 10 Then
            Label1.Text = "incorrect"
            Label1.ForeColor = Color.Red
            TextBox6.Focus()
            Label1.Show()
            Button1.Enabled = False
            Button5.Enabled = False
            Button6.Enabled = False
        Else
            Label1.Text = "correct"
            Label1.ForeColor = Color.Green
            Label1.Show()
            Button1.Enabled = True
            Button5.Enabled = True
            Button6.Enabled = True
        End If
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged

    End Sub

    Private Sub Panel9_Paint(sender As Object, e As PaintEventArgs) Handles Panel9.Paint

    End Sub
End Class