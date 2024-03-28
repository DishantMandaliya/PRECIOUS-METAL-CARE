Imports System.Data.OleDb
Imports System.Configuration
Public Class Additionalcharges
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim dt As New DataTable
    Public zaq1wsx As New DataTable
    Dim ds As New DataSet
    Dim x As New AutoCompleteStringCollection
    Dim c As Double
    Function total()
        c = Val(TextBox4.Text) * Val(TextBox2.Text)
        Return True
    End Function
    Private Sub Button11_Click(sender As Object, e As EventArgs) Handles Button11.Click
        Try
            zaq1wsx.Rows.Add(Label6.Text, Label7.Text, TextBox1.Text, TextBox4.Text, TextBox2.Text, c)
            sales.gb.Rows.Add("Additional charges", TextBox1.Text, TextBox4.Text, TextBox2.Text, c)
            DataGridView1.DataSource = zaq1wsx
            Label18.Text = 0
            For i = 0 To zaq1wsx.Rows.Count - 1
                Label18.Text = Val(Label18.Text) + Val(zaq1wsx.Rows(i).Item(5))
            Next
            TextBox1.Clear()
            TextBox2.Clear()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles DataGridView1.CellClick
        Try
            Label6.Text = zaq1wsx.Rows(DataGridView1.CurrentRow.Index).Item(0)
            TextBox1.Text = zaq1wsx.Rows(DataGridView1.CurrentRow.Index).Item(3)
            TextBox2.Text = zaq1wsx.Rows(DataGridView1.CurrentRow.Index).Item(4)
            TextBox4.Text = zaq1wsx.Rows(DataGridView1.CurrentRow.Index).Item(5)
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub Additionalcharges_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        TextBox2.Enabled = False
        TextBox4.Enabled = False
        zaq1wsx.Columns.Add("Add_ID")
        zaq1wsx.Columns.Add("sales_date1")
        zaq1wsx.Columns.Add("particulars_ps1")
        zaq1wsx.Columns.Add("quantity_qs1")
        zaq1wsx.Columns.Add("price_ps1")
        zaq1wsx.Columns.Add("Total_price1")
        TextBox4.Text = "01"
        Label6.Text = sales.Label1.Text
        Label7.Text = sales.DateTimePicker1.Text
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Me.Hide()
    End Sub
    Private Sub Additionalcharges_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        sales.Show()
    End Sub
    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        total()
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        total()
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        TextBox2.Enabled = True
        TextBox4.Enabled = True
    End Sub
End Class