Imports System.Data.OleDb
Imports TxtBoxControl

Public Class Fr_CN_Comers
    Dim da As New OleDbDataAdapter
    Dim Dttable As New DataTable
    Public NC As Integer

    Sub Clear_All()
        TextBox1.Text = ""
        TextBox2.Text = ""
        TextBox3.Text = ""
    End Sub

    Sub LD()
        Dttable.Clear()
        da = New OleDbDataAdapter("Select * from Comers where ID <> 6", Con)
        da.Fill(Dttable)
        DataGridView1.DataSource = Dttable
    End Sub

    Sub False_En()
        TextBox1.Enabled = False
        TextBox2.Enabled = False
        TextBox3.Enabled = False
        Button7.Enabled = True
        Button8.Enabled = True
    End Sub

    Sub True_En()
        TextBox1.Enabled = True
        TextBox2.Enabled = True
        TextBox3.Enabled = True
        Button7.Enabled = True
        Button8.Enabled = True
    End Sub

    Sub conn()
        TextBox1.DataBindings.Add("text", Dttable, "ID")
        TextBox2.DataBindings.Add("text", Dttable, "Comer_Name1")
        TextBox3.DataBindings.Add("text", Dttable, "Num_VAT")
    End Sub

    Sub cutt()
        TextBox2.DataBindings.Clear()
        TextBox3.DataBindings.Clear()
        TextBox1.DataBindings.Clear()
    End Sub

    Private Sub Fr_CN_Comers_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            LD()
            TextBox6.Text = "Start"
            Dim last As Integer = BindingContext(Dttable).Count - 1
            BindingContext(Dttable).Position = last
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub TextBox3_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TextBox3.KeyPress
        Try
            EnterNumber(sender, e, TextBox3, 25)
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            BindingContext(Dttable).Position = 0
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Try
            BindingContext(Dttable).Position -= 1
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Try
            BindingContext(Dttable).Position += 1
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Try
            Dim last As Integer = BindingContext(Dttable).Count - 1
            BindingContext(Dttable).Position = last
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        Try
            If TextBox6.Text = "Edit" Or TextBox6.Text = "Add" Then
                cutt()
            ElseIf TextBox6.Text = "Start" Or TextBox6.Text = "Start1" Then
                conn()
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        Try
            TextBox6.Text = "Add"
            Clear_All()
            True_En()
            Button5.Enabled = False
            Button6.Enabled = False
            Button1.Enabled = False
            Button2.Enabled = False
            Button3.Enabled = False
            Button4.Enabled = False
            Button9.Enabled = False
            TextBox2.Select()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button6_Click(sender As Object, e As EventArgs) Handles Button6.Click
        Try
            If TextBox2.Text = "" Then
                MsgBox("!لم يتم تحديد السجل المطلوب تعديله")
            Else
                True_En()
                TextBox6.Text = "Edit"
                Button1.Enabled = False
                Button2.Enabled = False
                Button3.Enabled = False
                Button4.Enabled = False
                Button5.Enabled = False
                Button6.Enabled = False
                Button9.Enabled = False
                TextBox2.Select()
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
                NC = TextBox1.Text
                Delete_Come(NC)
                MsgBox("تم الحذف بنجاح", MessageBoxButtons.OK, "تم الحذف")
                LD()
                NC = Nothing
                Close()
                Home_Form.TextBox1.Text = "open Come"
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button7_Click(sender As Object, e As EventArgs) Handles Button7.Click
        Try
            If TextBox2.Text <> "" And TextBox3.Text <> "" Then
                If TextBox6.Text = "Add" Then
                    Dttable.Rows.Add()
                    Dim Postion As Integer
                    Postion = Dttable.Rows.Count - 1
                    Insert_Come(TextBox2.Text, TextBox3.Text)
                End If
                If TextBox6.Text = "Edit" Then
                    NC = TextBox1.Text
                    Update_Come(TextBox2.Text, TextBox3.Text)
                    NC = Nothing
                End If
                MsgBox("تم الحفظ بنجاح", MessageBoxButtons.OK, "تم الحفظ")
                LD()
                TextBox6.Text = ""
                Close()
                Home_Form.TextBox1.Text = "open Come"
            Else
                Error_Form.MdiParent = Home_Form
                Error_Form.Show()
                Error_Form.TextBox1.Text = "Come"
                Me.Enabled = False
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Button8_Click(sender As Object, e As EventArgs) Handles Button8.Click
        Try
            If TextBox6.Text = "Add" Or TextBox6.Text = "Edit" Then
                Dim last As Integer = BindingContext(Dttable).Count - 1
                BindingContext(Dttable).Position = last
                False_En()
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
                Home_Form.TextBox1.Text = "open Come"
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub
End Class