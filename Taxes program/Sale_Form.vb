Imports System.Data.OleDb
Imports TxtBoxControl
Public Class Sale_Form
    Dim da As New OleDbDataAdapter
    Dim dtTable1 As New DataTable
    Dim Save As New OleDbCommandBuilder
    Dim Postion As Integer

    Sub changes()
        If TextBox6.Text = "Add" Then
            dtTable1.Rows.Add()
            Postion = dtTable1.Rows.Count - 1
        End If
        If TextBox6.Text = "Edit" Then
            Postion = BindingContext(dtTable1).Position
        End If
    End Sub

    Sub true_en()
        DateTimePicker1.Enabled = True
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox4.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
    End Sub

    Sub false_en()
        DateTimePicker1.Enabled = False
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox4.Enabled = False
    End Sub

    Sub The_Table()
        dtTable1.Clear()
        da = New OleDbDataAdapter("select * from The_Table where Type like 'مبيعات'", Con)
        da.Fill(dtTable1)
        DataGridView1.DataSource = dtTable1
    End Sub

    Sub Save_Table()
        Save = New OleDbCommandBuilder(da)
        da.Update(dtTable1)
        dtTable1.AcceptChanges()
    End Sub

    Sub Clear_All()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
        TextBox4.Text = ""
        TextBox5.Text = ""
    End Sub

    Sub Connection1()
        TextBox1.DataBindings.Add("text", dtTable1, "Num")
        TextBox4.DataBindings.Add("text", dtTable1, "Describe1")
        TextBox2.DataBindings.Add("text", dtTable1, "Price")
        TextBox3.DataBindings.Add("text", dtTable1, "vat")
        TextBox5.DataBindings.Add("text", dtTable1, "Type")
        DateTimePicker1.DataBindings.Add("value", dtTable1, "Date1")
    End Sub

    Sub Cut1()
        TextBox1.DataBindings.Clear()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox4.DataBindings.Clear()
        TextBox5.DataBindings.Clear()
        DateTimePicker1.DataBindings.Clear()
    End Sub

    Private Sub Sale_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            The_Table()
            TextBox6.Text = "Start"
            Dim last As Integer = BindingContext(dtTable1).Count - 1
            BindingContext(dtTable1).Position = last
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            BindingContext(dtTable1).Position = 0
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            BindingContext(dtTable1).Position -= 1
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            BindingContext(dtTable1).Position += 1
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim last As Integer = BindingContext(dtTable1).Count - 1
            BindingContext(dtTable1).Position = last
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            If TextBox1.Text <> "" And TextBox2.Text <> "" And TextBox3.Text <> "" Then
                changes()
                dtTable1.Rows(Postion).Item("Num") = TextBox1.Text
                dtTable1.Rows(Postion).Item("Date1") = DateTimePicker1.Value.ToShortDateString
                dtTable1.Rows(Postion).Item("Describe1") = TextBox4.Text
                dtTable1.Rows(Postion).Item("Price") = TextBox2.Text
                dtTable1.Rows(Postion).Item("vat") = TextBox3.Text
                dtTable1.Rows(Postion).Item("Type") = "مبيعات"
                dtTable1.Rows(Postion).Item("Comer_Name") = 6
                Save_Table()
                MsgBox("تم الحفظ بنجاح", MessageBoxButtons.OK, "تم الحفظ")
                The_Table()
                Close()
                Home_Form.TextBox1.Text = "open"
            Else
                Error_Form.MdiParent = Home_Form
                Error_Form.Show()
                Error_Form.TextBox1.Text = "Sale"
                Me.Enabled = False
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            TextBox6.Text = "Add"
            Clear_All()
            true_en()
            DateTimePicker1.Select()
            Button5.Enabled = False
            Button6.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button9.Enabled = False
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            If TextBox6.Text = "Add" Or TextBox6.Text = "Edit" Then
                Dim last As Integer = BindingContext(dtTable1).Count - 1
                BindingContext(dtTable1).Position = last
                false_en()
                Button7.Enabled = False
                Button8.Enabled = False
                Button5.Enabled = True
                Button6.Enabled = True
                Button1.Enabled = True
                Button2.Enabled = True
                Button3.Enabled = True
                Button4.Enabled = True
                Button9.Enabled = True
                Close()
                Home_Form.TextBox1.Text = "open"
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Try
            If TextBox6.Text = "Edit" Or TextBox6.Text = "Add" Then
                Cut1()
            ElseIf TextBox6.Text = "Start" Or TextBox6.Text = "Start1" Then
                Connection1()
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged
        Try
            If TextBox2.Text <> "" Then
                TextBox3.Text = TextBox2.Text / 20
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            If TextBox2.Text = "" Then
                MsgBox("!لم يتم تحديد السجل المطلوب تعديله")
            Else
                true_en()
                TextBox6.Text = "Edit"
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button5.Enabled = False
                Button6.Enabled = False
                Button9.Enabled = False
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button9_Click(sender As Object, e As EventArgs) Handles Button9.Click
        Try
            If TextBox2.Text = "" Then
                MsgBox("!لم يتم تحديد السجل المطلوب حذفه")
            Else
                Postion = BindingContext(dtTable1).Position
                dtTable1.Rows(Postion).Delete()
                Save_Table()
                MsgBox("تم الحذف بنجاح", MessageBoxButtons.OK, "تم الحذف")
                The_Table()
                Close()
                Home_Form.TextBox1.Text = "open"
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub TextBox1_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox1.KeyPress
        Try
            EnterNumber(sender, e, TextBox1, 7)
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub TextBox2_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox2.KeyPress
        If (e.KeyChar < "0" OrElse e.KeyChar > "9") _
            AndAlso e.KeyChar <> ControlChars.Back AndAlso e.KeyChar <> "." Then
            e.Handled = True
        End If
    End Sub
End Class