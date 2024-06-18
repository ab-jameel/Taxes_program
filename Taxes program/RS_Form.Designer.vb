<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class RS_Form
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim DataGridViewCellStyle1 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(RS_Form))
        Me.Button1 = New System.Windows.Forms.Button()
        Me.DataGridView1 = New System.Windows.Forms.DataGridView()
        Me.DateTimePicker1 = New System.Windows.Forms.DateTimePicker()
        Me.DateTimePicker2 = New System.Windows.Forms.DateTimePicker()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Button1
        '
        Me.Button1.Cursor = System.Windows.Forms.Cursors.Hand
        Me.Button1.Font = New System.Drawing.Font("Tahoma", 15.0!)
        Me.Button1.Location = New System.Drawing.Point(65, 14)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(127, 45)
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
        Me.DataGridView1.Location = New System.Drawing.Point(12, 76)
        Me.DataGridView1.MultiSelect = False
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.ReadOnly = True
        Me.DataGridView1.RowHeadersWidthSizeMode = System.Windows.Forms.DataGridViewRowHeadersWidthSizeMode.DisableResizing
        Me.DataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.DataGridView1.Size = New System.Drawing.Size(776, 428)
        Me.DataGridView1.TabIndex = 1
        '
        'DateTimePicker1
        '
        Me.DateTimePicker1.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.DateTimePicker1.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker1.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker1.Location = New System.Drawing.Point(348, 28)
        Me.DateTimePicker1.Name = "DateTimePicker1"
        Me.DateTimePicker1.Size = New System.Drawing.Size(148, 20)
        Me.DateTimePicker1.TabIndex = 1
        Me.DateTimePicker1.Value = New Date(2018, 1, 1, 0, 0, 0, 0)
        '
        'DateTimePicker2
        '
        Me.DateTimePicker2.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.DateTimePicker2.CustomFormat = "dd/MM/yyyy"
        Me.DateTimePicker2.Format = System.Windows.Forms.DateTimePickerFormat.Custom
        Me.DateTimePicker2.Location = New System.Drawing.Point(600, 28)
        Me.DateTimePicker2.Name = "DateTimePicker2"
        Me.DateTimePicker2.Size = New System.Drawing.Size(148, 20)
        Me.DateTimePicker2.TabIndex = 2
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label1.Location = New System.Drawing.Point(276, 29)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 17)
        Me.Label1.TabIndex = 4
        Me.Label1.Text = "من تاريخ :"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Tahoma", 10.0!)
        Me.Label2.Location = New System.Drawing.Point(528, 29)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(69, 17)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "الى تاريخ :"
        '
        'RS_Form
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(800, 516)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.DateTimePicker2)
        Me.Controls.Add(Me.DateTimePicker1)
        Me.Controls.Add(Me.DataGridView1)
        Me.Controls.Add(Me.Button1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "RS_Form"
        Me.RightToLeft = System.Windows.Forms.RightToLeft.Yes
        Me.RightToLeftLayout = True
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "التقرير الضريبي للمبيعات"
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Button1 As Button
    Friend WithEvents DataGridView1 As DataGridView
    Friend WithEvents DateTimePicker1 As DateTimePicker
    Friend WithEvents DateTimePicker2 As DateTimePicker
    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
End Class
