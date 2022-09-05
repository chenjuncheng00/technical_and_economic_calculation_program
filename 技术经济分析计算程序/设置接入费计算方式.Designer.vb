<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 设置接入费计算方式
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
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.清空窗体 = New System.Windows.Forms.Button()
        Me.默认方式 = New System.Windows.Forms.Button()
        Me.确定参数 = New System.Windows.Forms.Button()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.结束年份列表 = New System.Windows.Forms.ListBox()
        Me.开始年份列表 = New System.Windows.Forms.ListBox()
        Me.清空输入 = New System.Windows.Forms.Button()
        Me.添加输入 = New System.Windows.Forms.Button()
        Me.结束年份tmp = New System.Windows.Forms.TextBox()
        Me.开始年份tmp = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(305, 80)
        Me.Label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(229, 56)
        Me.Label4.TabIndex = 65
        Me.Label4.Text = "接入费结束年份"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(45, 80)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(233, 56)
        Me.Label3.TabIndex = 64
        Me.Label3.Text = "接入费开始年份"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(36, 25)
        Me.Label8.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(518, 50)
        Me.Label8.TabIndex = 70
        Me.Label8.Text = "设置接入费收入的开始年份和结束年份"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(590, 95)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(632, 32)
        Me.CheckBox1.TabIndex = 71
        Me.CheckBox1.Text = "接入费收入逐年计算比例是否乘以逐年达产率？"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        '清空窗体
        '
        Me.清空窗体.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空窗体.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空窗体.Location = New System.Drawing.Point(1118, 496)
        Me.清空窗体.Margin = New System.Windows.Forms.Padding(6)
        Me.清空窗体.Name = "清空窗体"
        Me.清空窗体.Size = New System.Drawing.Size(207, 82)
        Me.清空窗体.TabIndex = 74
        Me.清空窗体.Text = "清空窗体"
        Me.清空窗体.UseVisualStyleBackColor = False
        '
        '默认方式
        '
        Me.默认方式.BackColor = System.Drawing.SystemColors.ControlLight
        Me.默认方式.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.默认方式.Location = New System.Drawing.Point(862, 496)
        Me.默认方式.Margin = New System.Windows.Forms.Padding(6)
        Me.默认方式.Name = "默认方式"
        Me.默认方式.Size = New System.Drawing.Size(207, 82)
        Me.默认方式.TabIndex = 73
        Me.默认方式.Text = "默认方式"
        Me.默认方式.UseVisualStyleBackColor = False
        '
        '确定参数
        '
        Me.确定参数.BackColor = System.Drawing.SystemColors.ControlLight
        Me.确定参数.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.确定参数.Location = New System.Drawing.Point(604, 496)
        Me.确定参数.Margin = New System.Windows.Forms.Padding(6)
        Me.确定参数.Name = "确定参数"
        Me.确定参数.Size = New System.Drawing.Size(207, 82)
        Me.确定参数.TabIndex = 72
        Me.确定参数.Text = "确定参数"
        Me.确定参数.UseVisualStyleBackColor = False
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(590, 143)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(660, 32)
        Me.CheckBox2.TabIndex = 75
        Me.CheckBox2.Text = "接入费收入逐年计算比例是否乘以投产月份比例？"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(598, 192)
        Me.Label9.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(739, 50)
        Me.Label9.TabIndex = 102
        Me.Label9.Text = "接入费收入逐年负荷率计算结果显示"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(590, 258)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(747, 216)
        Me.RichTextBox1.TabIndex = 101
        Me.RichTextBox1.Text = ""
        '
        '结束年份列表
        '
        Me.结束年份列表.FormattingEnabled = True
        Me.结束年份列表.ItemHeight = 24
        Me.结束年份列表.Location = New System.Drawing.Point(348, 143)
        Me.结束年份列表.Name = "结束年份列表"
        Me.结束年份列表.Size = New System.Drawing.Size(142, 268)
        Me.结束年份列表.TabIndex = 108
        '
        '开始年份列表
        '
        Me.开始年份列表.FormattingEnabled = True
        Me.开始年份列表.ItemHeight = 24
        Me.开始年份列表.Location = New System.Drawing.Point(83, 143)
        Me.开始年份列表.Name = "开始年份列表"
        Me.开始年份列表.Size = New System.Drawing.Size(142, 268)
        Me.开始年份列表.TabIndex = 107
        '
        '清空输入
        '
        Me.清空输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空输入.Location = New System.Drawing.Point(301, 496)
        Me.清空输入.Margin = New System.Windows.Forms.Padding(6)
        Me.清空输入.Name = "清空输入"
        Me.清空输入.Size = New System.Drawing.Size(233, 82)
        Me.清空输入.TabIndex = 106
        Me.清空输入.Text = "清空输入"
        Me.清空输入.UseVisualStyleBackColor = False
        '
        '添加输入
        '
        Me.添加输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.添加输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.添加输入.Location = New System.Drawing.Point(41, 496)
        Me.添加输入.Margin = New System.Windows.Forms.Padding(6)
        Me.添加输入.Name = "添加输入"
        Me.添加输入.Size = New System.Drawing.Size(233, 82)
        Me.添加输入.TabIndex = 105
        Me.添加输入.Text = "添加输入"
        Me.添加输入.UseVisualStyleBackColor = False
        '
        '结束年份tmp
        '
        Me.结束年份tmp.Location = New System.Drawing.Point(348, 439)
        Me.结束年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.结束年份tmp.Name = "结束年份tmp"
        Me.结束年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.结束年份tmp.TabIndex = 104
        Me.结束年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '开始年份tmp
        '
        Me.开始年份tmp.Location = New System.Drawing.Point(83, 439)
        Me.开始年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.开始年份tmp.Name = "开始年份tmp"
        Me.开始年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.开始年份tmp.TabIndex = 103
        Me.开始年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(585, 25)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(730, 50)
        Me.Label1.TabIndex = 109
        Me.Label1.Text = "设置接入费收入的计算模式"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '设置接入费计算方式
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1372, 606)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.结束年份列表)
        Me.Controls.Add(Me.开始年份列表)
        Me.Controls.Add(Me.清空输入)
        Me.Controls.Add(Me.添加输入)
        Me.Controls.Add(Me.结束年份tmp)
        Me.Controls.Add(Me.开始年份tmp)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.清空窗体)
        Me.Controls.Add(Me.默认方式)
        Me.Controls.Add(Me.确定参数)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Name = "设置接入费计算方式"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置接入费计算方式"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents 清空窗体 As System.Windows.Forms.Button
    Friend WithEvents 默认方式 As System.Windows.Forms.Button
    Friend WithEvents 确定参数 As System.Windows.Forms.Button
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents 结束年份列表 As System.Windows.Forms.ListBox
    Friend WithEvents 开始年份列表 As System.Windows.Forms.ListBox
    Friend WithEvents 清空输入 As System.Windows.Forms.Button
    Friend WithEvents 添加输入 As System.Windows.Forms.Button
    Friend WithEvents 结束年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents 开始年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
