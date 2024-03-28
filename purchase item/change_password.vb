Imports System.Data.OleDb
Imports System.Configuration
Public Class change_password
    Dim conn As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim q As String
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        cmd = New OleDbCommand("select password from login where user_name = '" & TextBox3.Text & "'", conn)
        conn.Open()
        q = cmd.ExecuteScalar()
        'it send single value to the user from the database value of first row and first column it will send to the user. eg. count user name then answer will be 2 execute scalar will send 2 hence onyl a single value.
        If (q = TextBox1.Text) Then
            Label6.Text = "correct"
            Label6.ForeColor = Color.Green
        Else
            Label6.Text = "InCorrect"
            Label6.ForeColor = Color.Red
        End If
        conn.Close()
    End Sub
    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        If (TextBox4.Text = TextBox2.Text) Then
            Label7.Text = "Matched"
            Label7.ForeColor = Color.Green
            Button1.Enabled = True
        Else
            Label7.Text = "Not Matched"
            Label7.ForeColor = Color.Red
            Button1.Enabled = False
        End If
    End Sub
    Private Sub change_password_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Button1.Enabled = False
    End Sub
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Dim q As String
            q = "UPDATE login SET [Password] = '" & TextBox4.Text & "' WHERE [User_Name] = '" & TextBox3.Text & "'"
            cmd = New OleDbCommand(q, conn)
            cmd.Connection = conn
            conn.Open()
            cmd.ExecuteNonQuery()
            conn.Close()
            MsgBox("Updated Successfully!")
            clear()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.Source)
        End Try
    End Sub
    Function clear()
        TextBox1.Clear()
        TextBox2.Clear()
        TextBox3.Clear()
        TextBox4.Clear()
        Return True
    End Function
    Private Sub change_password_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Login.Show()
    End Sub
End Class