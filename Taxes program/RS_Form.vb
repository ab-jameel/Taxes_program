Imports System.Data.OleDb
Public Class RS_Form
    Dim da As New OleDbDataAdapter
    Dim dtTable1 As New DataTable

    Sub Search()
        Dim SearchDate1 As Date = DateTimePicker1.Value
        Dim SearchDate2 As Date = DateTimePicker2.Value.AddDays(1)
        dtTable1.Clear()
        da = New OleDbDataAdapter("select * from The_Table where Date1 > #" & SearchDate1.Year & "/" & SearchDate1.Month & "/" & SearchDate1.Day & "# and Date1 < #" & SearchDate2.Year & "/" & SearchDate2.Month & "/" & SearchDate2.Day & "# and Type like 'مبيعات'", Con)
        da.Fill(dtTable1)
        DataGridView1.DataSource = dtTable1
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "رقم الفاتورة"
        DataGridView1.Columns(2).HeaderText = "تاريخ الفاتورة"
        DataGridView1.Columns(3).HeaderText = "البيان"
        DataGridView1.Columns(4).HeaderText = "المبلغ"
        DataGridView1.Columns(5).HeaderText = "الضريبة"
        DataGridView1.Columns(6).HeaderText = "نوع الفاتورة"
        DataGridView1.Columns(6).Visible = False
        DataGridView1.Columns(2).DefaultCellStyle.Format = "dd/MM/yyyy"
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(3).Width = 332
    End Sub

    Private Sub RS_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Search()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Search()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub
End Class