Imports System.Data.OleDb
Imports System.Configuration
Public Class Login
    Dim conn As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
    Dim adp As New OleDbDataAdapter
    Dim cmd As New OleDbCommand
    Dim dr As OleDbDataReader
    Dim ds As New DataSet
    Dim q As String
    Dim dt123 As New DataTable
    Dim x As New AutoCompleteStringCollection
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            checkconnection()
            Dim q As String
            q = "select * from login where user_name = '" + TextBox2.Text + "' and Password = '" + TextBox1.Text + "'"
            cmd.Connection = conn
            cmd.CommandText = q
            conn.Open()
            dr = cmd.ExecuteReader()
            If (dr.Read) Then
                MsgBox("Login Succesfull")
                Me.Hide()
                Homepage.Show()
                conn.Close()
            Else
                MsgBox("Enter Proper Credentials")
            End If
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub
    Private Sub checkconnection()
        If (conn.State = ConnectionState.Open) Then
            conn.Close()
        End If
    End Sub
    Private Sub Login_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Dim r As String
            Dim dt As New DataTable
            r = "select user_name from login"
            Dim adp As New OleDbDataAdapter(r, conn)
            adp.Fill(dt)
            suggest()
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
        Button1.Enabled = False
    End Sub
    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Button1.Enabled = True
    End Sub
    Function suggest()
        Try
            q = "select User_Name from login"
            Dim adp As New OleDbDataAdapter(q, conn)
            adp.Fill(dt123)
            For i = 0 To dt123.Rows.Count - 1
                x.Add(dt123.Rows(i)("User_Name").ToString)
            Next
            TextBox2.AutoCompleteMode = AutoCompleteMode.Suggest
            TextBox2.AutoCompleteSource = AutoCompleteSource.CustomSource
            TextBox2.AutoCompleteCustomSource = x
        Catch ex As Exception
            MsgBox(ex.Message)
            MsgBox(ex.StackTrace)
        End Try
        Return True
    End Function
    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        change_password.Show()
    End Sub
End Class