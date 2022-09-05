<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 设置补贴收入计算年限
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
        Me.Label3 = New System.Windows.Forms.Label()
        Me.结束年份tmp = New System.Windows.Forms.TextBox()
        Me.开始年份tmp = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.补贴收入1 = New System.Windows.Forms.Button()
        Me.补贴收入2 = New System.Windows.Forms.Button()
        Me.清空窗体 = New System.Windows.Forms.Button()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.光伏补贴收入 = New System.Windows.Forms.Button()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.添加输入 = New System.Windows.Forms.Button()
        Me.清空输入 = New System.Windows.Forms.Button()
        Me.开始年份列表 = New System.Windows.Forms.ListBox()
        Me.结束年份列表 = New System.Windows.Forms.ListBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(35, 28)
        Me.Label8.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(500, 50)
        Me.Label8.TabIndex = 44
        Me.Label8.Text = "设置补贴收入的开始年份和结束年份"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(51, 82)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(194, 56)
        Me.Label3.TabIndex = 38
        Me.Label3.Text = "补贴开始年份"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '结束年份tmp
        '
        Me.结束年份tmp.Location = New System.Drawing.Point(347, 584)
        Me.结束年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.结束年份tmp.Name = "结束年份tmp"
        Me.结束年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.结束年份tmp.TabIndex = 37
        Me.结束年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '开始年份tmp
        '
        Me.开始年份tmp.Location = New System.Drawing.Point(73, 584)
        Me.开始年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.开始年份tmp.Name = "开始年份tmp"
        Me.开始年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.开始年份tmp.TabIndex = 32
        Me.开始年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(323, 82)
        Me.Label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(194, 56)
        Me.Label4.TabIndex = 53
        Me.Label4.Text = "补贴结束年份"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '补贴收入1
        '
        Me.补贴收入1.BackColor = System.Drawing.SystemColors.ControlLight
        Me.补贴收入1.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.补贴收入1.Location = New System.Drawing.Point(575, 548)
        Me.补贴收入1.Margin = New System.Windows.Forms.Padding(6)
        Me.补贴收入1.Name = "补贴收入1"
        Me.补贴收入1.Size = New System.Drawing.Size(341, 82)
        Me.补贴收入1.TabIndex = 62
        Me.补贴收入1.Text = "补贴收入1"
        Me.补贴收入1.UseVisualStyleBackColor = False
        '
        '补贴收入2
        '
        Me.补贴收入2.BackColor = System.Drawing.SystemColors.ControlLight
        Me.补贴收入2.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.补贴收入2.Location = New System.Drawing.Point(942, 548)
        Me.补贴收入2.Margin = New System.Windows.Forms.Padding(6)
        Me.补贴收入2.Name = "补贴收入2"
        Me.补贴收入2.Size = New System.Drawing.Size(341, 82)
        Me.补贴收入2.TabIndex = 63
        Me.补贴收入2.Text = "补贴收入2"
        Me.补贴收入2.UseVisualStyleBackColor = False
        '
        '清空窗体
        '
        Me.清空窗体.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空窗体.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空窗体.Location = New System.Drawing.Point(942, 643)
        Me.清空窗体.Margin = New System.Windows.Forms.Padding(6)
        Me.清空窗体.Name = "清空窗体"
        Me.清空窗体.Size = New System.Drawing.Size(341, 82)
        Me.清空窗体.TabIndex = 64
        Me.清空窗体.Text = "清空窗体内容"
        Me.清空窗体.UseVisualStyleBackColor = False
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(588, 94)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(604, 32)
        Me.CheckBox1.TabIndex = 66
        Me.CheckBox1.Text = "补贴收入逐年计算比例是否乘以逐年达产率？"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(588, 148)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(688, 32)
        Me.CheckBox2.TabIndex = 67
        Me.CheckBox2.Text = "补贴收入逐年计算比例是否乘以逐年达产率增加值？"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        '光伏补贴收入
        '
        Me.光伏补贴收入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.光伏补贴收入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.光伏补贴收入.Location = New System.Drawing.Point(575, 643)
        Me.光伏补贴收入.Margin = New System.Windows.Forms.Padding(6)
        Me.光伏补贴收入.Name = "光伏补贴收入"
        Me.光伏补贴收入.Size = New System.Drawing.Size(341, 82)
        Me.光伏补贴收入.TabIndex = 68
        Me.光伏补贴收入.Text = "光伏补贴收入"
        Me.光伏补贴收入.UseVisualStyleBackColor = False
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(588, 201)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(660, 32)
        Me.CheckBox3.TabIndex = 69
        Me.CheckBox3.Text = "补贴收入逐年计算比例是否乘以投产月份数比例？"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(575, 306)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(708, 216)
        Me.RichTextBox1.TabIndex = 70
        Me.RichTextBox1.Text = ""
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(575, 241)
        Me.Label9.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(707, 50)
        Me.Label9.TabIndex = 71
        Me.Label9.Text = "补贴收入逐年负荷率计算结果显示"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '添加输入
        '
        Me.添加输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.添加输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.添加输入.Location = New System.Drawing.Point(40, 643)
        Me.添加输入.Margin = New System.Windows.Forms.Padding(6)
        Me.添加输入.Name = "添加输入"
        Me.添加输入.Size = New System.Drawing.Size(233, 82)
        Me.添加输入.TabIndex = 72
        Me.添加输入.Text = "添加输入"
        Me.添加输入.UseVisualStyleBackColor = False
        '
        '清空输入
        '
        Me.清空输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空输入.Location = New System.Drawing.Point(302, 643)
        Me.清空输入.Margin = New System.Windows.Forms.Padding(6)
        Me.清空输入.Name = "清空输入"
        Me.清空输入.Size = New System.Drawing.Size(233, 82)
        Me.清空输入.TabIndex = 73
        Me.清空输入.Text = "清空输入"
        Me.清空输入.UseVisualStyleBackColor = False
        '
        '开始年份列表
        '
        Me.开始年份列表.FormattingEnabled = True
        Me.开始年份列表.ItemHeight = 24
        Me.开始年份列表.Location = New System.Drawing.Point(73, 147)
        Me.开始年份列表.Name = "开始年份列表"
        Me.开始年份列表.Size = New System.Drawing.Size(142, 412)
        Me.开始年份列表.TabIndex = 74
        '
        '结束年份列表
        '
        Me.结束年份列表.FormattingEnabled = True
        Me.结束年份列表.ItemHeight = 24
        Me.结束年份列表.Location = New System.Drawing.Point(347, 147)
        Me.结束年份列表.Name = "结束年份列表"
        Me.结束年份列表.Size = New System.Drawing.Size(142, 412)
        Me.结束年份列表.TabIndex = 75
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(575, 28)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(708, 50)
        Me.Label1.TabIndex = 76
        Me.Label1.Text = "设置补贴收入的计算模式"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '设置补贴收入计算年限
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1320, 756)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.结束年份列表)
        Me.Controls.Add(Me.开始年份列表)
        Me.Controls.Add(Me.清空输入)
        Me.Controls.Add(Me.添加输入)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.光伏补贴收入)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.清空窗体)
        Me.Controls.Add(Me.补贴收入2)
        Me.Controls.Add(Me.补贴收入1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.结束年份tmp)
        Me.Controls.Add(Me.开始年份tmp)
        Me.Margin = New System.Windows.Forms.Padding(6)
        Me.Name = "设置补贴收入计算年限"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置补贴收入计算年限"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents 结束年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents 开始年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents 补贴收入1 As System.Windows.Forms.Button
    Friend WithEvents 补贴收入2 As System.Windows.Forms.Button
    Friend WithEvents 清空窗体 As System.Windows.Forms.Button
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents 光伏补贴收入 As System.Windows.Forms.Button
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents 添加输入 As System.Windows.Forms.Button
    Friend WithEvents 清空输入 As System.Windows.Forms.Button
    Friend WithEvents 开始年份列表 As System.Windows.Forms.ListBox
    Friend WithEvents 结束年份列表 As System.Windows.Forms.ListBox
    Friend WithEvents Label1 As System.Windows.Forms.Label
End Class
