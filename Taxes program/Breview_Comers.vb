Imports System.Data.OleDb
Public Class Breview_Comers
    Dim da5 As New OleDbDataAdapter
    Dim dtTable5 As New DataTable

    Sub Comers_Table()
        dtTable5.Clear()
        da5 = New OleDbDataAdapter("select * from Comers where ID <> 6", Con)
        da5.Fill(dtTable5)
        DataGridView1.DataSource = dtTable5
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).Width = 135
        DataGridView1.Columns(2).Width = 175
        DataGridView1.Columns(1).HeaderText = "اسم المورد"
        DataGridView1.Columns(2).HeaderText = "رقم السجل الضريبي"
    End Sub

    Private Sub Breview_Comers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Comers_Table()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub DataGridView1_DoubleClick(sender As Object, e As EventArgs) Handles DataGridView1.DoubleClick
        If DataGridView1.Rows.Count > 0 Then
            Try
                Buy_Form.NComer.Text = Me.DataGridView1.CurrentRow.Cells(1).Value
                Buy_Form.NumVat.Text = Me.DataGridView1.CurrentRow.Cells(2).Value
                Buy_Form.IDComer.Text = Me.DataGridView1.CurrentRow.Cells(0).Value
                Me.Close()
                Buy_Form.Enabled = True
            Catch
                MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
            End Try
        Else
            MsgBox("الرجاء اختيار المورد",, "لم يتم الاختيار")
        End If
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            dtTable5.Clear()
        da5 = New OleDbDataAdapter("Select * from Comers where Comer_Name1 like '%" & TextBox1.Text.Trim & "%' and ID <> 6", Con)
        da5.Fill(dtTable5)
        DataGridView1.DataSource = dtTable5
        DataGridView1.Columns(0).Visible = False
        DataGridView1.Columns(1).Width = 135
        DataGridView1.Columns(2).Width = 175
        DataGridView1.Columns(1).HeaderText = "اسم المورد"
            DataGridView1.Columns(2).HeaderText = "رقم السجل الضريبي"
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub
End Class