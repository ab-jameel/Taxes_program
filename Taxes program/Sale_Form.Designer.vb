<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Sale_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(Sale_Form))
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.TextBox2 = New System.Windows.Forms.TextBox()
        Me.TextBox3 = New System.Windows.Forms.TextBox()
        Me.TextBox4 = New System.Windows.Forms.TextBox()
        Me.TextBox5 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Button5 = New System.Windows.Forms.Button()
        Me.Button6 = New System.Windows.Forms.Button()
        Me.Button7 = New System.Windows.Forms.Button()
        Me.Button8 = New System.Windows.Forms.Button()
        Me.TextBox6 = New System.Windows.Forms.TextBox()
        Me.Button9 = New System.Windows.Forms.Button()
        Me.Button4 = New System.Windows.Forms.Button()
        Me.Button3 = New System.Windows.Forms.Button()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Location = New System.Drawing.Point(350, 89)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(100, 19)
        Me.DataGridView1.TabIndex = 0
        Me.DataGridView1.Visible = False
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Enabled = False
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(166, 34)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.DateTimePicker1.RightToLeftLayout = True
        Me.DateTimePicker1.Size = New System.Drawing.Size(153, 20)
        Me.DateTimePicker1.TabIndex = 1
        Me.DateTimePicker1.Value = New Date(2018, 9, 12, 0, 0, 0, 0)
        '
        'TextBox1
        '
        Me.TextBox1.Enabled = False
        Me.TextBox1.Location = New System.Drawing.Point(525, 34)
        Me.TextBox1.MaxLength = 7
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(163, 20)
        Me.TextBox1.TabIndex = 2
        '
        'TextBox2
        '
        Me.TextBox2.Enabled = False
        Me.TextBox2.Location = New System.Drawing.Point(166, 83)
        Me.TextBox2.MaxLength = 6
        Me.TextBox2.Name = "TextBox2"
        Me.TextBox2.Size = New System.Drawing.Size(153, 20)
        Me.TextBox2.TabIndex = 3
        '
        'TextBox3
        '
        Me.TextBox3.Enabled = False
        Me.TextBox3.Location = New System.Drawing.Point(525, 85)
        Me.TextBox3.Name = "TextBox3"
        Me.TextBox3.Size = New System.Drawing.Size(163, 20)
        Me.TextBox3.TabIndex = 4
        '
        'TextBox4
        '
        Me.TextBox4.Enabled = False
        Me.TextBox4.Location = New System.Drawing.Point(29, 161)
        Me.TextBox4.MaxLength = 250
        Me.TextBox4.Multiline = True
        Me.TextBox4.Name = "TextBox4"
        Me.TextBox4.Size = New System.Drawing.Size(659, 91)
        Me.TextBox4.TabIndex = 5
        '
        'TextBox5
        '
        Me.TextBox5.Location = New System.Drawing.Point(350, 68)
        Me.TextBox5.Name = "TextBox5"
        Me.TextBox5.Size = New System.Drawing.Size(100, 20)
        Me.TextBox5.TabIndex = 6
        Me.TextBox5.Visible = False
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label1.Location = New System.Drawing.Point(26, 34)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(132, 17)
        Me.Label1.TabIndex = 7
        Me.Label1.Text = "تاريخ فاتورة المبيعات :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label2.Location = New System.Drawing.Point(393, 34)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(126, 17)
        Me.Label2.TabIndex = 8
        Me.Label2.Text = "رقم فاتورة المبيعات :"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label3.Location = New System.Drawing.Point(107, 86)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(51, 17)
        Me.Label3.TabIndex = 9
        Me.Label3.Text = "المبلغ :"
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label4.Location = New System.Drawing.Point(461, 86)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(58, 17)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "الضريبة :"
        '
        'Label5
        '
        Me.Label5.AutoSize = True
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label5.Location = New System.Drawing.Point(26, 141)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(47, 17)
        Me.Label5.TabIndex = 11
        Me.Label5.Text = "البيان :"
        '
        'Button5
        '
        Me.Button5.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button5.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Button5.Location = New System.Drawing.Point(29, 269)
        Me.Button5.Name = "Button5"
        Me.Button5.Size = New System.Drawing.Size(127, 50)
        Me.Button5.TabIndex = 6
        Me.Button5.Text = "جديد"
        Me.Button5.UseVisualStyleBackColor = True
        '
        'Button6
        '
        Me.Button6.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button6.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Button6.Location = New System.Drawing.Point(162, 269)
        Me.Button6.Name = "Button6"
        Me.Button6.Size = New System.Drawing.Size(127, 50)
        Me.Button6.TabIndex = 7
        Me.Button6.Text = "تعديل"
        Me.Button6.UseVisualStyleBackColor = True
        '
        'Button7
        '
        Me.Button7.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button7.Enabled = False
        Me.Button7.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Button7.Location = New System.Drawing.Point(428, 269)
        Me.Button7.Name = "Button7"
        Me.Button7.Size = New System.Drawing.Size(127, 50)
        Me.Button7.TabIndex = 9
        Me.Button7.Text = "حفظ"
        Me.Button7.UseVisualStyleBackColor = True
        '
        'Button8
        '
        Me.Button8.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button8.Enabled = False
        Me.Button8.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Button8.Location = New System.Drawing.Point(561, 269)
        Me.Button8.Name = "Button8"
        Me.Button8.Size = New System.Drawing.Size(127, 50)
        Me.Button8.TabIndex = 10
        Me.Button8.Text = "الغاء"
        Me.Button8.UseVisualStyleBackColor = True
        '
        'TextBox6
        '
        Me.TextBox6.Location = New System.Drawing.Point(350, 114)
        Me.TextBox6.Name = "TextBox6"
        Me.TextBox6.Size = New System.Drawing.Size(100, 20)
        Me.TextBox6.TabIndex = 20
        Me.TextBox6.Visible = False
        '
        'Button9
        '
        Me.Button9.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button9.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Button9.Location = New System.Drawing.Point(295, 269)
        Me.Button9.Name = "Button9"
        Me.Button9.Size = New System.Drawing.Size(127, 50)
        Me.Button9.TabIndex = 8
        Me.Button9.Text = "حذف"
        Me.Button9.UseVisualStyleBackColor = True
        '
        'Button4
        '
        Me.Button4.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button4.Image = CType(resources.GetObject("Button4.Image"), System.Drawing.Image)
        Me.Button4.Location = New System.Drawing.Point(421, 349)
        Me.Button4.Name = "Button4"
        Me.Button4.Size = New System.Drawing.Size(71, 30)
        Me.Button4.TabIndex = 14
        Me.Button4.UseVisualStyleBackColor = True
        '
        'Button3
        '
        Me.Button3.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button3.Image = CType(resources.GetObject("Button3.Image"), System.Drawing.Image)
        Me.Button3.Location = New System.Drawing.Point(344, 349)
        Me.Button3.Name = "Button3"
        Me.Button3.Size = New System.Drawing.Size(71, 30)
        Me.Button3.TabIndex = 13
        Me.Button3.UseVisualStyleBackColor = True
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Image = CType(resources.GetObject("Button2.Image"), System.Drawing.Image)
        Me.Button2.Location = New System.Drawing.Point(267, 349)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(71, 30)
        Me.Button2.TabIndex = 12
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Image = CType(resources.GetObject("Button1.Image"), System.Drawing.Image)
        Me.Button1.Location = New System.Drawing.Point(190, 349)
        Me.Button1.Name = "Button1"
        Me.Button1.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.Button1.Size = New System.Drawing.Size(71, 30)
        Me.Button1.TabIndex = 11
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Sale_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(737, 397)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Button3)
        Me.Controls.Add(Me.Button9)
        Me.Controls.Add(Me.Button4)
        Me.Controls.Add(Me.TextBox6)
        Me.Controls.Add(Me.Button8)
        Me.Controls.Add(Me.Button7)
        Me.Controls.Add(Me.Button6)
        Me.Controls.Add(Me.Button5)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.TextBox5)
        Me.Controls.Add(Me.TextBox4)
        Me.Controls.Add(Me.TextBox3)
        Me.Controls.Add(Me.TextBox2)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.DataGridView1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "Sale_Form"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "حساب ضريبة فواتير المبيعات"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents TextBox4 As TextBox
    Friend WithEvents TextBox5 As TextBox
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Button5 As Button
    Friend WithEvents Button6 As Button
    Friend WithEvents Button7 As Button
    Friend WithEvents Button8 As Button
    Friend WithEvents TextBox6 As TextBox
    Friend WithEvents Button9 As Button
    Friend WithEvents Button4 As Button
    Friend WithEvents Button3 As Button
    Friend WithEvents Button2 As Button
    Friend WithEvents Button1 As Button
End Class
