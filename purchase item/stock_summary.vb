Public Class stock_summary
    Dim datat As New DataTable
    Private Sub stock_summary_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        sea("SELECT particulars.particular_name, Count(*) AS stock, Sum(stock_table.Net_Weight) AS Total_Weight FROM particulars INNER JOIN stock_table ON particulars.ID = stock_table.particulars_id GROUP BY particulars.particular_name, stock_table.particulars_id;", datat)
        DataGridView1.DataSource = datat
        RadioButton3.Checked = True
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        check()
        Label2.Text = "Total Weight(Gold) :- "
        tot()
    End Sub
    Function check()
        If RadioButton1.Checked = True Then
            sea("SELECT particulars.particular_name, Count(*) AS stock, Sum(stock_table.Net_Weight) AS Total_Weight FROM particulars INNER JOIN stock_table ON particulars.ID = stock_table.particulars_id WHERE(((stock_table.category_id) = 1))GROUP BY particulars.particular_name, stock_table.particulars_id;", datat)
            DataGridView1.DataSource = datat
        ElseIf RadioButton2.Checked = True Then
            sea("SELECT particulars.particular_name, Count(*) AS stock, Sum(stock_table.Net_Weight) AS Total_Weight FROM particulars INNER JOIN stock_table ON particulars.ID = stock_table.particulars_id WHERE(((stock_table.category_id) = 2))GROUP BY particulars.particular_name, stock_table.particulars_id;", datat)
            DataGridView1.DataSource = datat
        Else
            sea("SELECT particulars.particular_name, Count(*) AS stock, Sum(stock_table.Net_Weight) AS Total_Weight FROM particulars INNER JOIN stock_table ON particulars.ID = stock_table.particulars_id GROUP BY particulars.particular_name, stock_table.particulars_id;", datat)
            DataGridView1.DataSource = datat
        End If
        Return True
    End Function
    Function tot()
        Label1.Show()
        Label2.Show()
        Label1.Text = 0
        For i = 0 To datat.Rows.Count - 1
            Label1.Text = Val(Label1.Text) + Val(datat.Rows(i).Item(2))
        Next
        Return True
    End Function

    Private Sub RadioButton2_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton2.CheckedChanged
        check()
        Label2.Text = "Total Weight(Silver) :- "
        tot()
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        check()
        Label1.Hide()
        Label2.Hide()
    End Sub
End Class