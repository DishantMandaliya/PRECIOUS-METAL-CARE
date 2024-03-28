Imports System.Data.OleDb
Imports System.Configuration
Public Class search
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim dt, zaq1 As New DataTable
    Dim ds As New DataSet
    Dim x As New AutoCompleteStringCollection
    Dim c As Double

    Public Function search(q As String)
        adp = New OleDbDataAdapter(q, comm)
        Try
            zaq1.Clear()
            adp.Fill(zaq1)
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
        Return True
    End Function
    Public Function showthing()
        If ComboBox1.SelectedIndex = 0 Then
            'search1.Label3.Text = "Date"
            'search1.Label5.Text = "particular"
            search1.Label6.Hide()
            search1.TextBox6.Hide()
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        ElseIf ComboBox1.SelectedIndex = 1 Then
            'search1.Label3.Text = "Date"
            'search1.Label5.Text = "particular"
            search1.Label6.Hide()
            search1.TextBox6.Hide()
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        ElseIf ComboBox1.SelectedIndex = 2 Then
            'search1.Label3.Text = "Date"
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        ElseIf ComboBox1.SelectedIndex = 3 Then
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        ElseIf ComboBox1.SelectedIndex = 4 Then
            'search1.Label3.Text = "Address"
            'search1.Label4.Text = "Phone Number"
            'search1.Label5.Text = "References"
            search1.Label6.Hide()
            search1.TextBox6.Hide()
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        ElseIf ComboBox1.SelectedIndex = 5 Then
            'search1.Label3.Text = "Address"
            'search1.Label4.Text = "Phone Number"
            'search1.Label5.Text = "City"
            'search1.Label6.Text = "State"
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        ElseIf ComboBox1.SelectedIndex = 6 Then
            search1.TextBox5.Show()
            search1.Label5.Show()
            search1.Label5.Text = "particular"
            'search1.Label3.Text = "Particular"
            search1.Label2.Text = "Category"
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()

        Else
            search1.Label1.Text = "particular"
            search1.Label2.Text = "category"
            search1.Label3.Text = "purity value"
            search1.Label4.Text = "purity_vs"
            search1.Label6.Hide()
            search1.TextBox6.Hide()
        End If
        Return True
    End Function

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        'search click==========
        search1.Button1.Hide()
        If ComboBox1.SelectedIndex = 7 Then
            Label1.Show()
            Label1.Text = "You Can't Search On Others Please Select different Catagory"
        Else
            Me.Hide()
            search1.MdiParent = Homepage
            search1.Show()
            search1.Dock = DockStyle.Fill
            search1.Button1.Text = "Search"

            search1.TextBox2.Hide()
            search1.Label3.Hide()
            search1.TextBox3.Hide()
            search1.Label4.Hide()

            search1.Label5.Hide()
            search1.TextBox5.Hide()
            search1.Label6.Hide()
            search1.TextBox6.Hide()
        End If
        showthing()
    End Sub

    Private Sub search_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Hide()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        '========add==========
        Me.Hide()
        If ComboBox1.SelectedIndex = 0 Then
            Me.Hide()
            purchase.MdiParent = Homepage
            purchase.Dock = DockStyle.Fill
            purchase.Show()
        ElseIf ComboBox1.SelectedIndex = 1 Then
            Me.Hide()
            sales.MdiParent = Homepage
            sales.Dock = DockStyle.Fill
            sales.Show()
        ElseIf ComboBox1.SelectedIndex = 2 Then
            Me.Hide()
            order_Book.MdiParent = Homepage
            order_Book.Dock = DockStyle.Fill
            order_Book.Show()
        ElseIf ComboBox1.SelectedIndex = 3 Then
            Me.Hide()
            payment.MdiParent = Homepage
            payment.Show()
        ElseIf ComboBox1.SelectedIndex = 4 Then
            Me.Hide()
            customer.MdiParent = Homepage
            customer.Show()
        ElseIf ComboBox1.SelectedIndex = 5 Then
            Me.Hide()
            supplier.MdiParent = Homepage
            supplier.Show()
        ElseIf ComboBox1.SelectedIndex = 6 Then
            Me.Hide()
            stock_details.MdiParent = Homepage
            stock_details.Dock = DockStyle.Fill
            stock_details.Show()
        Else
            Me.Hide()
            search1.MdiParent = Homepage
            search1.Show()
            search1.Dock = DockStyle.Fill
        End If
        search1.Button1.Text = "Add"
        showthing()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        '===========delete===========
        If ComboBox1.SelectedIndex = 4 Then
            Me.Hide()
            customer.MdiParent = Homepage
            customer.Show()
        ElseIf ComboBox1.SelectedIndex = 5 Then
            Me.Hide()
            supplier.MdiParent = Homepage
            supplier.Show()
        ElseIf ComboBox1.SelectedIndex = 6 Then
            Me.Hide()
            stock_details.MdiParent = Homepage
            stock_details.Dock = DockStyle.Fill
            stock_details.Show()
        ElseIf ComboBox1.SelectedIndex = 7 Then
            Me.Hide()
            search1.MdiParent = Homepage
            search1.Show()
            search1.Dock = DockStyle.Fill
            search1.RadioButton1.Show()
            search1.RadioButton2.Show()
            search1.RadioButton3.Show()
            search1.RadioButton4.Show()
        Else
            Me.Hide()
            search1.MdiParent = Homepage
            search1.Show()
            search1.Dock = DockStyle.Fill
            search1.RadioButton1.Hide()
            search1.RadioButton2.Hide()
            search1.RadioButton3.Hide()
            search1.RadioButton4.Hide()
        End If
        search1.Button1.Text = "Delete"
        search1.TextBox1.Enabled = True
        search1.Label2.Hide()
        search1.TextBox2.Hide()
        search1.Label3.Hide()
        search1.TextBox3.Hide()
        search1.Label4.Hide()
        search1.TextBox4.Hide()
        search1.Label5.Hide()
        search1.TextBox5.Hide()
        search1.Label6.Hide()
        search1.TextBox6.Hide()

    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label1.Hide()
    End Sub
End Class