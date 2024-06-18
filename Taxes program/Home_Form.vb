Public Class Home_Form
    Sub Clear_All()
        RB_Form.Close()
        RS_Form.Close()
        RT_Form.Close()
        Buy_Form.Close()
        Sale_Form.Close()
        Fr_Comers.Close()
        Fr_CN_Comers.Close()
        Breview_Comers.Close()
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            If TextBox1.Text = "open" Then
                Clear_All()
                Sale_Form.MdiParent = Me
                Sale_Form.Show()
                Sale_Form.TextBox6.Text = "Start"
                TextBox1.Text = ""
            ElseIf TextBox1.Text = "open buy" Then
                Clear_All()
                Buy_Form.MdiParent = Me
                Buy_Form.Show()
                Buy_Form.TextBox6.Text = "Start"
                TextBox1.Text = ""
            ElseIf TextBox1.Text = "open Come" Then
                Clear_All()
                Fr_CN_Comers.MdiParent = Me
                Fr_CN_Comers.Show()
                Fr_CN_Comers.TextBox6.Text = "Start"
                TextBox1.Text = ""
            End If
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub Home_Form_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            For Each ctrl As Control In Me.Controls
                If TypeOf ctrl Is MdiClient Then
                    ctrl.BackgroundImage = Image.FromFile(Application.StartupPath & "\background.jpg")
                End If
            Next ctrl
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem5_Click(sender As Object, e As EventArgs) Handles ButtonItem5.Click
        Try
            Clear_All()
            Sale_Form.MdiParent = Me
            Sale_Form.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem4_Click(sender As Object, e As EventArgs) Handles ButtonItem4.Click
        Try
            Clear_All()
            Buy_Form.MdiParent = Me
            Buy_Form.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem3_Click(sender As Object, e As EventArgs) Handles ButtonItem3.Click
        Try
            Clear_All()
            RS_Form.MdiParent = Me
            RS_Form.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem2_Click(sender As Object, e As EventArgs) Handles ButtonItem2.Click
        Try
            Clear_All()
            RB_Form.MdiParent = Me
            RB_Form.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem1_Click(sender As Object, e As EventArgs) Handles ButtonItem1.Click
        Try
            Clear_All()
            RT_Form.MdiParent = Me
            RT_Form.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem19_Click(sender As Object, e As EventArgs) Handles ButtonItem19.Click
        Try
            Clear_All()
            Fr_Comers.MdiParent = Me
            Fr_Comers.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub

    Private Sub ButtonItem20_Click(sender As Object, e As EventArgs) Handles ButtonItem20.Click
        Try
            Clear_All()
            Fr_CN_Comers.MdiParent = Me
            Fr_CN_Comers.Show()
        Catch
            MsgBox("يوجد خطأ الرجاء التواصل مع الدعم الفني",, "خطأ")
        End Try
    End Sub
End Class