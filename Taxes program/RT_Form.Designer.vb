<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class RT_Form
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
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RT_Form))
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.SaveFileDialog1 = New System.Windows.Forms.SaveFileDialog()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label2.Location = New System.Drawing.Point(129, 51)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 17)
        Me.Label2.TabIndex = 15
        Me.Label2.Text = "الى تاريخ :"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label1.Location = New System.Drawing.Point(129, 13)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 17)
        Me.Label1.TabIndex = 14
        Me.Label1.Text = "من تاريخ :"
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(201, 11)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(114, 20)
        Me.DateTimePicker1.TabIndex = 1
        Me.DateTimePicker1.Value = New Date(2018, 1, 1, 0, 0, 0, 0)
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Location = New System.Drawing.Point(19, 5)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(93, 30)
        Me.Button1.TabIndex = 3
        Me.Button1.Text = "عرض"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'DataGridView1
        '
        Me.DataGridView1.AllowUserToAddRows = False
        Me.DataGridView1.AllowUserToDeleteRows = False
        Me.DataGridView1.AllowUserToResizeColumns = False
        Me.DataGridView1.AllowUserToResizeRows = False
        DataGridViewCellStyle1.BackColor = System.Drawing.Color.FromArgb(CType(CType(128, Byte), Integer), CType(CType(255, Byte), Integer), CType(CType(128, Byte), Integer))
        Me.DataGridView1.AlternatingRowsDefaultCellStyle = DataGridViewCellStyle1
        Me.DataGridView1.AutoSizeRowsMode = System.Windows.Forms.DataGridViewAutoSizeRowsMode.AllCells
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.Color.White
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Tahoma", 8.0!)
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.DataGridView1.DefaultCellStyle = DataGridViewCellStyle2
        Me.DataGridView1.Location = New System.Drawing.Point(12, 77)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(987, 428)
        Me.DataGridView1.TabIndex = 17
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label3.Location = New System.Drawing.Point(406, 8)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(201, 27)
        Me.Label3.TabIndex = 21
        Me.Label3.Text = "اجمالي ضرائب فواتير المبيعات :"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label4.Location = New System.Drawing.Point(390, 41)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(217, 27)
        Me.Label4.TabIndex = 22
        Me.Label4.Text = "اجمالي ضرائب فواتير المشتريات :"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("Tahoma", 11.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label5.Location = New System.Drawing.Point(755, 20)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(91, 41)
        Me.Label5.TabIndex = 23
        Me.Label5.Text = "الرصيد :"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.BackColor = System.Drawing.Color.White
        Me.Label6.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label6.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(613, 8)
        Me.Label6.Name = "Label6"
        Me.Label6.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label6.Size = New System.Drawing.Size(110, 27)
        Me.Label6.TabIndex = 24
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.BackColor = System.Drawing.Color.White
        Me.Label7.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label7.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label7.Location = New System.Drawing.Point(613, 41)
        Me.Label7.Name = "Label7"
        Me.Label7.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label7.Size = New System.Drawing.Size(110, 27)
        Me.Label7.TabIndex = 25
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label8.Font = New System.Drawing.Font("Tahoma", 20.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label8.Location = New System.Drawing.Point(852, 8)
        Me.Label8.Name = "Label8"
        Me.Label8.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label8.Size = New System.Drawing.Size(110, 60)
        Me.Label8.TabIndex = 26
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label9
        '
        Me.Label9.BackColor = System.Drawing.Color.White
        Me.Label9.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label9.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label9.Location = New System.Drawing.Point(446, 199)
        Me.Label9.Name = "Label9"
        Me.Label9.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label9.Size = New System.Drawing.Size(110, 27)
        Me.Label9.TabIndex = 27
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label9.Visible = False
        '
        'Label10
        '
        Me.Label10.BackColor = System.Drawing.Color.White
        Me.Label10.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D
        Me.Label10.Font = New System.Drawing.Font("Tahoma", 14.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label10.Location = New System.Drawing.Point(450, 291)
        Me.Label10.Name = "Label10"
        Me.Label10.RightToLeft = System.Windows.Forms.RightToLeft.No
        Me.Label10.Size = New System.Drawing.Size(110, 27)
        Me.Label10.TabIndex = 28
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        Me.Label10.Visible = False
        '
        'Button2
        '
        Me.Button2.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button2.Location = New System.Drawing.Point(19, 37)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(93, 37)
        Me.Button2.TabIndex = 4
        Me.Button2.Text = "تحويل الى ملف excel"
        Me.Button2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.Button2.UseVisualStyleBackColor = True
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(201, 51)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(114, 20)
        Me.DateTimePicker2.TabIndex = 2
        '
        'RT_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1011, 516)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "RT_Form"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "الاقرار الضريبي"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label2 As Label
    Friend WithEvents Label1 As Label
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents TextBox2 As TextBox
    Friend WithEvents TextBox3 As TextBox
    Friend WithEvents Label3 As Label
    Friend WithEvents Label4 As Label
    Friend WithEvents Label5 As Label
    Friend WithEvents Label6 As Label
    Friend WithEvents Label7 As Label
    Friend WithEvents Label8 As Label
    Friend WithEvents Label9 As Label
    Friend WithEvents Label10 As Label
    Friend WithEvents Button2 As Button
    Friend WithEvents SaveFileDialog1 As SaveFileDialog
    Friend WithEvents DateTimePicker2 As DateTimePicker
End Class
