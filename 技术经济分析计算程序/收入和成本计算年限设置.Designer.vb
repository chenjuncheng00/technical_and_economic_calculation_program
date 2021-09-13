<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 收入和成本计算年限设置
    Inherits System.Windows.Forms.Form

    'Form 重写 Dispose，以清理组件列表。
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

    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer

    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.jsnf1 = New System.Windows.Forms.TextBox()
        Me.jsnf2 = New System.Windows.Forms.TextBox()
        Me.jsnf3 = New System.Windows.Forms.TextBox()
        Me.ksnf3 = New System.Windows.Forms.TextBox()
        Me.ksnf2 = New System.Windows.Forms.TextBox()
        Me.ksnf1 = New System.Windows.Forms.TextBox()
        Me.购电容量费成本 = New System.Windows.Forms.Button()
        Me.城市管廊成本 = New System.Windows.Forms.Button()
        Me.充电桩收入 = New System.Windows.Forms.Button()
        Me.人员工资 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.人员工资递增比例 = New System.Windows.Forms.TextBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(6, 34)
        Me.Label8.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(828, 50)
        Me.Label8.TabIndex = 87
        Me.Label8.Text = "设置收入和成本的开始年份和结束年份，最多设置3段"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(426, 253)
        Me.Label5.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(237, 56)
        Me.Label5.TabIndex = 86
        Me.Label5.Text = "收入成本结束年份："
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(9, 253)
        Me.Label6.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(242, 56)
        Me.Label6.TabIndex = 85
        Me.Label6.Text = "收入成本开始年份："
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(426, 175)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(237, 56)
        Me.Label1.TabIndex = 84
        Me.Label1.Text = "收入成本结束年份："
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(7, 175)
        Me.Label2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(244, 56)
        Me.Label2.TabIndex = 83
        Me.Label2.Text = "收入成本开始年份："
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(426, 103)
        Me.Label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(237, 56)
        Me.Label4.TabIndex = 82
        Me.Label4.Text = "收入成本结束年份："
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(11, 103)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(240, 56)
        Me.Label3.TabIndex = 81
        Me.Label3.Text = "收入成本开始年份："
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'jsnf1
        '
        Me.jsnf1.Location = New System.Drawing.Point(664, 109)
        Me.jsnf1.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf1.Name = "jsnf1"
        Me.jsnf1.Size = New System.Drawing.Size(142, 35)
        Me.jsnf1.TabIndex = 80
        Me.jsnf1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsnf2
        '
        Me.jsnf2.Location = New System.Drawing.Point(664, 189)
        Me.jsnf2.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf2.Name = "jsnf2"
        Me.jsnf2.Size = New System.Drawing.Size(142, 35)
        Me.jsnf2.TabIndex = 79
        Me.jsnf2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsnf3
        '
        Me.jsnf3.Location = New System.Drawing.Point(664, 263)
        Me.jsnf3.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf3.Name = "jsnf3"
        Me.jsnf3.Size = New System.Drawing.Size(142, 35)
        Me.jsnf3.TabIndex = 78
        Me.jsnf3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksnf3
        '
        Me.ksnf3.Location = New System.Drawing.Point(256, 263)
        Me.ksnf3.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf3.Name = "ksnf3"
        Me.ksnf3.Size = New System.Drawing.Size(142, 35)
        Me.ksnf3.TabIndex = 77
        Me.ksnf3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksnf2
        '
        Me.ksnf2.Location = New System.Drawing.Point(256, 189)
        Me.ksnf2.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf2.Name = "ksnf2"
        Me.ksnf2.Size = New System.Drawing.Size(142, 35)
        Me.ksnf2.TabIndex = 76
        Me.ksnf2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksnf1
        '
        Me.ksnf1.Location = New System.Drawing.Point(256, 109)
        Me.ksnf1.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf1.Name = "ksnf1"
        Me.ksnf1.Size = New System.Drawing.Size(142, 35)
        Me.ksnf1.TabIndex = 75
        Me.ksnf1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '购电容量费成本
        '
        Me.购电容量费成本.BackColor = System.Drawing.SystemColors.ControlLight
        Me.购电容量费成本.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.购电容量费成本.Location = New System.Drawing.Point(57, 878)
        Me.购电容量费成本.Margin = New System.Windows.Forms.Padding(6)
        Me.购电容量费成本.Name = "购电容量费成本"
        Me.购电容量费成本.Size = New System.Drawing.Size(341, 82)
        Me.购电容量费成本.TabIndex = 92
        Me.购电容量费成本.Text = "购电容量费成本"
        Me.购电容量费成本.UseVisualStyleBackColor = False
        '
        '城市管廊成本
        '
        Me.城市管廊成本.BackColor = System.Drawing.SystemColors.ControlLight
        Me.城市管廊成本.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.城市管廊成本.Location = New System.Drawing.Point(430, 780)
        Me.城市管廊成本.Margin = New System.Windows.Forms.Padding(6)
        Me.城市管廊成本.Name = "城市管廊成本"
        Me.城市管廊成本.Size = New System.Drawing.Size(341, 82)
        Me.城市管廊成本.TabIndex = 91
        Me.城市管廊成本.Text = "城市管廊成本"
        Me.城市管廊成本.UseVisualStyleBackColor = False
        '
        '充电桩收入
        '
        Me.充电桩收入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.充电桩收入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.充电桩收入.Location = New System.Drawing.Point(57, 780)
        Me.充电桩收入.Margin = New System.Windows.Forms.Padding(6)
        Me.充电桩收入.Name = "充电桩收入"
        Me.充电桩收入.Size = New System.Drawing.Size(341, 82)
        Me.充电桩收入.TabIndex = 89
        Me.充电桩收入.Text = "充电桩收入"
        Me.充电桩收入.UseVisualStyleBackColor = False
        '
        '人员工资
        '
        Me.人员工资.BackColor = System.Drawing.SystemColors.ControlLight
        Me.人员工资.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.人员工资.Location = New System.Drawing.Point(430, 878)
        Me.人员工资.Margin = New System.Windows.Forms.Padding(6)
        Me.人员工资.Name = "人员工资"
        Me.人员工资.Size = New System.Drawing.Size(341, 82)
        Me.人员工资.TabIndex = 94
        Me.人员工资.Text = "人员工资"
        Me.人员工资.UseVisualStyleBackColor = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(25, 328)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(352, 32)
        Me.CheckBox1.TabIndex = 95
        Me.CheckBox1.Text = "人员工资是否逐年递增？"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(365, 315)
        Me.Label7.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(298, 56)
        Me.Label7.TabIndex = 97
        Me.Label7.Text = "人员工资逐年递增比例(%)"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '人员工资递增比例
        '
        Me.人员工资递增比例.Location = New System.Drawing.Point(664, 325)
        Me.人员工资递增比例.Margin = New System.Windows.Forms.Padding(6)
        Me.人员工资递增比例.Name = "人员工资递增比例"
        Me.人员工资递增比例.Size = New System.Drawing.Size(142, 35)
        Me.人员工资递增比例.TabIndex = 96
        Me.人员工资递增比例.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(25, 384)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(688, 32)
        Me.CheckBox2.TabIndex = 98
        Me.CheckBox2.Text = "收入和成本逐年计算系数是否乘以投产月份数比例？"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(57, 479)
        Me.Label9.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(713, 50)
        Me.Label9.TabIndex = 100
        Me.Label9.Text = "部分收入和成本逐年负荷率计算结果显示"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(57, 542)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(713, 216)
        Me.RichTextBox1.TabIndex = 99
        Me.RichTextBox1.Text = ""
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(25, 437)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(688, 32)
        Me.CheckBox3.TabIndex = 101
        Me.CheckBox3.Text = "充电桩收入逐年计算系数是否乘以逐年负荷达产率？"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        '收入和成本计算年限设置
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(834, 983)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.人员工资递增比例)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.人员工资)
        Me.Controls.Add(Me.购电容量费成本)
        Me.Controls.Add(Me.城市管廊成本)
        Me.Controls.Add(Me.充电桩收入)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.jsnf1)
        Me.Controls.Add(Me.jsnf2)
        Me.Controls.Add(Me.jsnf3)
        Me.Controls.Add(Me.ksnf3)
        Me.Controls.Add(Me.ksnf2)
        Me.Controls.Add(Me.ksnf1)
        Me.Name = "收入和成本计算年限设置"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置部分销售收入和经营成本计算年限"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents jsnf1 As System.Windows.Forms.TextBox
    Friend WithEvents jsnf2 As System.Windows.Forms.TextBox
    Friend WithEvents jsnf3 As System.Windows.Forms.TextBox
    Friend WithEvents ksnf3 As System.Windows.Forms.TextBox
    Friend WithEvents ksnf2 As System.Windows.Forms.TextBox
    Friend WithEvents ksnf1 As System.Windows.Forms.TextBox
    Friend WithEvents 购电容量费成本 As System.Windows.Forms.Button
    Friend WithEvents 城市管廊成本 As System.Windows.Forms.Button
    Friend WithEvents 充电桩收入 As System.Windows.Forms.Button
    Friend WithEvents 人员工资 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents 人员工资递增比例 As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
End Class
