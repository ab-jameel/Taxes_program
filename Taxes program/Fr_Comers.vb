Imports System.Data.OleDb
Public Class Fr_Comers
    Dim da As New OleDbDataAdapter
    Dim DtTable As New DataTable
    Private Sub Fr_Comers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            DtTable.Clear()
            da = New OleDbDataAdapter("Select * from Comers where ID <> 6", Con)
            da.Fill(DtTable)
            DataGridView1.DataSource = DtTable
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).HeaderText = "اسم المورد"
            DataGridView1.Columns(2).HeaderText = "رقم السجل الضريبي"
            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(2).Width = 294
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            DtTable.Clear()
            da = New OleDbDataAdapter("Select * from Comers where Comer_Name1 like '%" & TextBox1.Text.Trim & "%' and ID <> 6", Con)
            da.Fill(DtTable)
            DataGridView1.DataSource = DtTable
            DataGridView1.Columns(0).Visible = False
            DataGridView1.Columns(1).HeaderText = "اسم المورد"
            DataGridView1.Columns(2).HeaderText = "رقم السجل الضريبي"
            DataGridView1.Columns(1).Width = 200
            DataGridView1.Columns(2).Width = 294
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub
End Class