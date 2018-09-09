Imports System.Data.OleDb
Public Class Form1
    Dim conn As New OleDbConnection
    Dim dbProvider As String = "PROVIDER=Microsoft.Jet.OLEDB.4.0;"
    Dim dbSource As String = "Data Source = D:\AccessDB\db_vbtodo.mdb;"
    Dim adapter As OleDbDataAdapter
    Dim ds As DataSet
    Dim currentid As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        conn.ConnectionString = dbProvider & dbSource
        GetRecords()
    End Sub

    Public Sub GetRecords()
        ds = New DataSet

        conn.Open()
        adapter = New OleDbDataAdapter("select * from [tbl_tasks]", conn)
        adapter.Fill(ds, "tbl_tasks")

        DataGridView1.DataSource = ds
        DataGridView1.DataMember = "tbl_tasks"
        conn.Close()
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        ds = New DataSet
        adapter = New OleDbDataAdapter("insert into [tbl_tasks] ([task_name]) VALUES ('" & TextBox1.Text & "')", conn)
        adapter.Fill(ds, "tbl_tasks")
        TextBox1.Clear()
        GetRecords()
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        currentid = DataGridView1.Item(0, i).Value.ToString()

        ds = New DataSet
        adapter = New OleDbDataAdapter("delete from [tbl_tasks] where id = " & currentid, conn)
        adapter.Fill(ds, "tbl_tasks")

        GetRecords()
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Dim i As Integer
        i = DataGridView1.CurrentRow.Index
        currentid = DataGridView1.Item(0, i).Value.ToString()
        TextBox1.Text = DataGridView1.Item(1, i).Value.ToString()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ds = New DataSet
        adapter = New OleDbDataAdapter("update [tbl_tasks] set task_name = '" & TextBox1.Text &
                                       "' where id = " & currentid, conn)
        adapter.Fill(ds, "tbl_tasks")
        TextBox1.Clear()
        GetRecords()
    End Sub
End Class
