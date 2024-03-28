
Imports System.Data.OleDb
Imports System.Configuration
Module load_data

    Public dt10, dt11, dt12, dt13, dt14, dt15, dt21 As New DataTable

    Public Function datashow()
        Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
        Dim q As String
        Dim adp10, adp11, adp12, adp13 As New OleDbDataAdapter
        dt10.Clear()
        dt11.Clear()
        dt12.Clear()
        dt13.Clear()
        q = "select * from category"
        adp10 = New OleDbDataAdapter(q, comm)
        adp10.Fill(dt10)

        q = "select * from purity"
        adp11 = New OleDbDataAdapter(q, comm)
        adp11.Fill(dt11)

        q = "select * from particulars"
        adp12 = New OleDbDataAdapter(q, comm)
        adp12.Fill(dt12)

        q = "select * from js_size"
        adp13 = New OleDbDataAdapter(q, comm)
        adp13.Fill(dt13)
        Return True
    End Function

    Public Function state()
        Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
        Dim q As String
        Dim adp15 As New OleDbDataAdapter

        dt15.Clear()
        q = "select * from state_select"
        adp15 = New OleDbDataAdapter(q, comm)
        adp15.Fill(dt15)

        Return True
    End Function
    Public Function sea(q As String, dt21 As DataTable)
        Dim comm As New OleDbConnection(ConfigurationManager.ConnectionStrings("conn").ConnectionString)
        Dim adp21 As New OleDbDataAdapter(q, comm)

        Try
            dt21.Clear()
            adp21.Fill(dt21)

        Catch ex As Exception
            MsgBox(ex.Message)

        End Try
        Return True
    End Function
End Module
