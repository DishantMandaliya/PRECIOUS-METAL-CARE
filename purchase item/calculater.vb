Public Class calculater
    Dim a, b, c, d As Double
    Private Sub calculater_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Label4.Text = ""
        Label2.Text = ""
        Label3.Enabled = False
        TextBox3.Enabled = False
        Button2.Enabled = False
        Label4.Enabled = False
    End Sub
    Function calc()
        Try
            a = TextBox1.Text
            b = TextBox2.Text
            c = (a * b) / 100
            Label2.Text = c
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
        Return True
    End Function
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        calc()
        Label3.Enabled = True
        TextBox3.Enabled = True
        Button2.Enabled = True
        Label4.Enabled = True
    End Sub
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If RadioButton2.Checked = True Then
            Label4.Text = (TextBox3.Text * Label2.Text) / 1000
        Else
            Label4.Text = (TextBox3.Text * Label2.Text) / 10
        End If
    End Sub
    Function clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        Label4.Text = ""
        Label2.Text = ""
        Return True
    End Function
    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        clear()
    End Sub
End Class