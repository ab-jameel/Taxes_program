Public Class Error_Form
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Try
            Me.Close()
        If TextBox1.Text = "Sale" Then
            Sale_Form.Enabled = True
        ElseIf TextBox1.Text = "Buy" Then
            Buy_Form.Enabled = True
        ElseIf TextBox1.Text = "Come" Then
            Fr_CN_Comers.Enabled = True
        End If
            TextBox1.Text = ""
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub
End Class