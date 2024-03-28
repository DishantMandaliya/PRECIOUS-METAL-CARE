Imports System.Data.OleDb
Imports System.Configuration
Public Class Billings
    Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim d, dt As New DataTable
    Dim ds As New DataSet
    Dim cmd As New OleDbCommand
    Dim zaq1 As New DataTable
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            PrintPreviewDialog1.Document = PrintDocument1
            PrintPreviewDialog1.PrintPreviewControl.Zoom = 1
            PrintPreviewDialog1.ShowDialog()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Private Sub Billings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label6.Text = sales.ComboBox6.Text
        Label8.Text = sales.DateTimePicker1.Text
        Label3.Text = sales.Label1.Text
        Label7.Text = sales.TextBox6.Text
        DataGridView1.DataSource = sales.gb
        For i = 0 To sales.gb.Rows.Count - 1
            Label4.Text = Val(Label4.Text) + Val(sales.gb.Rows(i).Item(4))
        Next
        PrintPreviewDialog1.Document = PrintDocument1
    End Sub
    Private Sub PrintDocument1_PrintPage(sender As Object, e As Printing.PrintPageEventArgs) Handles PrintDocument1.PrintPage
        Dim imgbnp As New Bitmap(Me.Width, Me.Height)
        DrawToBitmap(imgbnp, New Rectangle(0, 0, Me.Width, Me.Height))
        e.Graphics.DrawImage(imgbnp, 0, 0)
    End Sub
    Private Sub Billings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        sales.gb.Clear()
        Me.Hide()
        sales.Show()
    End Sub
End Class