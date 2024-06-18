Imports System.Data.OleDb
Public Class RB_Form
    Dim da As New OleDbDataAdapter
    Dim dtTable1 As New DataTable

    Sub Search()
        Dim SearchDate1 As Date = DateTimePicker1.Value
        Dim SearchDate2 As Date = DateTimePicker2.Value.AddDays(1)
        dtTable1.Clear()
        da = New OleDbDataAdapter("select The_Table.ID, The_Table.Num, The_Table.Date1, The_Table.Describe1,
                                  The_Table.Price, The_Table.VAT, The_Table.Type, The_Table.Comer_Name , 
                                  Comers.Comer_Name1,Comers.Num_VAT
                                  from The_Table,Comers 
                                  where The_Table.Type like 'مشتريات'
                                  and The_Table.Comer_Name = Comers.ID 
                                  and The_Table.Date1 > #" & SearchDate1.Year & "/" & SearchDate1.Month & "/" & SearchDate1.Day & "# and The_Table.Date1 < #" & SearchDate2.Year & "/" & SearchDate2.Month & "/" & SearchDate2.Day & "#", Con)
        da.Fill(dtTable1)
        DataGridView1.DataSource = dtTable1
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).HeaderText = "رقم الفاتورة"
        DataGridView1.Columns(2).HeaderText = "تاريخ الفاتورة"
        DataGridView1.Columns(3).HeaderText = "البيان"
        DataGridView1.Columns(4).HeaderText = "المبلغ"
        DataGridView1.Columns(5).HeaderText = "الضريبة"
        DataGridView1.Columns("Type").HeaderText = "نوع الفاتورة"
        DataGridView1.Columns(6).Visible = False
        DataGridView1.Columns(2).DefaultCellStyle.Format = "dd/MM/yyyy"
        DataGridView1.Columns(3).Width = 250
        DataGridView1.Columns(4).Width = 91
        DataGridView1.Columns(5).Width = 91
        DataGridView1.Columns(7).Visible = False
        DataGridView1.Columns(8).HeaderText = "اسم المورد"
        DataGridView1.Columns(9).HeaderText = "رقم السجل الضريبي"
        DataGridView1.Columns(9).Visible = False
        DataGridView1.Columns(8).Width = 101
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