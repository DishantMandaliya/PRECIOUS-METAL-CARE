Public Class Homepage
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        stock_summary.Hide()
        order_Book.Hide()
        calculater.Hide()
        sales.Hide()
        stock_details.Hide()
        order_Book.Hide()
        search1.Close()
        search.Close()
        supplier.Hide()
        payment.Hide()
        customer.Hide()
        purchase.MdiParent = Me
        purchase.Show()
        purchase.Dock = DockStyle.Fill
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        stock_summary.Hide()
        order_Book.Hide()
        calculater.Hide()
        search1.Close()
        search.Close()
        purchase.Hide()
        stock_details.Hide()
        order_Book.Hide()
        supplier.Hide()
        payment.Hide()
        customer.Hide()
        sales.MdiParent = Me
        sales.Show()
        sales.Dock = DockStyle.Fill
    End Sub
    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Me.Hide()
        Login.Show()
    End Sub
    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            ActiveMdiChild.Close()
        Catch
        End Try
    End Sub
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            stock_summary.Close()
            order_Book.Close()
            calculater.Hide()
            sales.Hide()
            purchase.Hide()
            order_Book.Hide()
            search1.Close()
            search.Close()
            supplier.Hide()
            payment.Hide()
            customer.Hide()
            stock_details.MdiParent = Me
            stock_details.Show()
            stock_details.Dock = DockStyle.Fill
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub Homepage_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label1.Text = "Namaste"
        Label2.Text = Login.TextBox2.Text
    End Sub
    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        stock_summary.Hide()
        order_Book.Close()
        calculater.Hide()
        sales.Hide()
        search1.Close()
        search.Close()
        purchase.Hide()
        stock_details.Hide()
        supplier.Hide()
        payment.Hide()
        customer.Hide()
        order_Book.MdiParent = Me
        order_Book.Show()
        order_Book.Dock = DockStyle.Fill
    End Sub
    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        stock_summary.Close()
        order_Book.Close()
        calculater.Close()
        search1.Close()
        sales.Close()
        purchase.Close()
        stock_details.Close()
        order_Book.Close()
        supplier.Close()
        payment.Close()
        customer.Close()
        search.MdiParent = Me
        search.Show()
    End Sub

    Private Sub ToolStripSplitButton1_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton1.ButtonClick
        calculater.StartPosition = FormStartPosition.CenterParent
        calculater.MdiParent = Me
        calculater.Show()
    End Sub

    Private Sub ToolStripSplitButton2_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton2.ButtonClick
        view_order.MdiParent = Me
        view_order.Show()
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        stock_summary.Hide()
        order_Book.Hide()
        calculater.Hide()
        search1.Close()
        search.Close()
        sales.Hide()
        purchase.Hide()
        stock_details.Hide()
        order_Book.Hide()
        supplier.Hide()
        customer.Hide()
        payment.MdiParent = Me
        payment.Show()
    End Sub
    Private Sub ToolStripSplitButton3_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton3.ButtonClick
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Button1.BackColor = ColorDialog1.Color
            Button2.BackColor = ColorDialog1.Color
            Button3.BackColor = ColorDialog1.Color
            ToolStripSplitButton3.BackColor = ColorDialog1.Color
            Button5.BackColor = ColorDialog1.Color
            Button6.BackColor = ColorDialog1.Color
            Button7.BackColor = ColorDialog1.Color
            Button8.BackColor = ColorDialog1.Color
            Button9.BackColor = ColorDialog1.Color
            Button4.BackColor = ColorDialog1.Color
            Button10.BackColor = ColorDialog1.Color
        End If
    End Sub
    Private Sub ToolStripSplitButton4_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton4.ButtonClick
        If ColorDialog1.ShowDialog() = Windows.Forms.DialogResult.OK Then
            Button1.ForeColor = ColorDialog1.Color
            Button2.ForeColor = ColorDialog1.Color
            Button3.ForeColor = ColorDialog1.Color
            Button4.ForeColor = ColorDialog1.Color
            ToolStripSplitButton4.BackColor = ColorDialog1.Color
            Button5.ForeColor = ColorDialog1.Color
            Button6.ForeColor = ColorDialog1.Color
            Button7.ForeColor = ColorDialog1.Color
            Button8.ForeColor = ColorDialog1.Color
            Button9.ForeColor = ColorDialog1.Color
            Button10.ForeColor = ColorDialog1.Color
        End If
    End Sub
    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        stock_summary.Hide()
        order_Book.Hide()
        calculater.Hide()
        sales.Hide()
        search1.Close()
        search.Close()
        purchase.Hide()
        stock_details.Hide()
        order_Book.Hide()
        supplier.Hide()
        payment.Hide()
        customer.MdiParent = Me
        customer.Show()
    End Sub
    Private Sub Button10_Click(sender As Object, e As EventArgs) Handles Button10.Click
        stock_summary.Hide()
        order_Book.Hide()
        calculater.Hide()
        search1.Close()
        search.Close()
        sales.Hide()
        purchase.Hide()
        stock_details.Hide()
        order_Book.Hide()
        customer.Hide()
        payment.Hide()
        supplier.MdiParent = Me
        supplier.Show()
    End Sub
    Private Sub ToolStripSplitButton5_ButtonClick(sender As Object, e As EventArgs) Handles ToolStripSplitButton5.ButtonClick
        stock_summary.MdiParent = Me
        stock_summary.Show()
        stock_summary.StartPosition = FormStartPosition.CenterScreen

    End Sub
End Class