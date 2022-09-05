<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class 设置逐年衰减计算方式
    Inherits System.Windows.Forms.Form
    'Form 重写 Dispose，以清理组件列表。
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
    'Windows 窗体设计器所必需的
    Private components As System.ComponentModel.IContainer
    '注意: 以下过程是 Windows 窗体设计器所必需的
    '可以使用 Windows 窗体设计器修改它。  
    '不要使用代码编辑器修改它。
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.开始年份tmp = New System.Windows.Forms.TextBox()
        Me.衰减率tmp = New System.Windows.Forms.TextBox()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.光伏发电 = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.蓄电池供电 = New System.Windows.Forms.Button()
        Me.清空窗体 = New System.Windows.Forms.Button()
        Me.光伏发电默认系数 = New System.Windows.Forms.Button()
        Me.蓄电池默认系数 = New System.Windows.Forms.Button()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.gfjsnf = New System.Windows.Forms.TextBox()
        Me.xdcjsnf = New System.Windows.Forms.TextBox()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.xdcksnf = New System.Windows.Forms.TextBox()
        Me.gfksnf = New System.Windows.Forms.TextBox()
        Me.CheckBox1 = New System.Windows.Forms.CheckBox()
        Me.蓄电池购电 = New System.Windows.Forms.Button()
        Me.清空输入 = New System.Windows.Forms.Button()
        Me.添加输入 = New System.Windows.Forms.Button()
        Me.衰减率列表 = New System.Windows.Forms.ListBox()
        Me.开始年份列表 = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        '开始年份tmp
        '
        Me.开始年份tmp.Location = New System.Drawing.Point(72, 599)
        Me.开始年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.开始年份tmp.Name = "开始年份tmp"
        Me.开始年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.开始年份tmp.TabIndex = 0
        Me.开始年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '衰减率tmp
        '
        Me.衰减率tmp.Location = New System.Drawing.Point(321, 599)
        Me.衰减率tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.衰减率tmp.Name = "衰减率tmp"
        Me.衰减率tmp.Size = New System.Drawing.Size(142, 35)
        Me.衰减率tmp.TabIndex = 5
        Me.衰减率tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(26, 79)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(231, 56)
        Me.Label3.TabIndex = 8
        Me.Label3.Text = "衰减开始年份"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(273, 72)
        Me.Label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(231, 56)
        Me.Label4.TabIndex = 11
        Me.Label4.Text = "逐年衰减率(%)"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(28, 22)
        Me.Label7.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(482, 50)
        Me.Label7.TabIndex = 14
        Me.Label7.Text = "输入开始衰减的年份以及逐年衰减率"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '光伏发电
        '
        Me.光伏发电.BackColor = System.Drawing.SystemColors.ControlLight
        Me.光伏发电.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.光伏发电.Location = New System.Drawing.Point(548, 556)
        Me.光伏发电.Margin = New System.Windows.Forms.Padding(6)
        Me.光伏发电.Name = "光伏发电"
        Me.光伏发电.Size = New System.Drawing.Size(202, 82)
        Me.光伏发电.TabIndex = 17
        Me.光伏发电.Text = "光伏发电"
        Me.光伏发电.UseVisualStyleBackColor = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(548, 315)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(717, 216)
        Me.RichTextBox1.TabIndex = 18
        Me.RichTextBox1.Text = ""
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(548, 253)
        Me.Label9.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(708, 50)
        Me.Label9.TabIndex = 19
        Me.Label9.Text = "光伏发电OR蓄电池逐年负荷率计算结果显示"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '蓄电池供电
        '
        Me.蓄电池供电.BackColor = System.Drawing.SystemColors.ControlLight
        Me.蓄电池供电.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.蓄电池供电.Location = New System.Drawing.Point(548, 650)
        Me.蓄电池供电.Margin = New System.Windows.Forms.Padding(6)
        Me.蓄电池供电.Name = "蓄电池供电"
        Me.蓄电池供电.Size = New System.Drawing.Size(206, 82)
        Me.蓄电池供电.TabIndex = 20
        Me.蓄电池供电.Text = "蓄电池供电"
        Me.蓄电池供电.UseVisualStyleBackColor = False
        '
        '清空窗体
        '
        Me.清空窗体.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空窗体.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空窗体.Location = New System.Drawing.Point(770, 556)
        Me.清空窗体.Margin = New System.Windows.Forms.Padding(6)
        Me.清空窗体.Name = "清空窗体"
        Me.清空窗体.Size = New System.Drawing.Size(206, 82)
        Me.清空窗体.TabIndex = 21
        Me.清空窗体.Text = "清空窗体"
        Me.清空窗体.UseVisualStyleBackColor = False
        '
        '光伏发电默认系数
        '
        Me.光伏发电默认系数.BackColor = System.Drawing.SystemColors.ControlLight
        Me.光伏发电默认系数.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.光伏发电默认系数.Location = New System.Drawing.Point(996, 556)
        Me.光伏发电默认系数.Margin = New System.Windows.Forms.Padding(6)
        Me.光伏发电默认系数.Name = "光伏发电默认系数"
        Me.光伏发电默认系数.Size = New System.Drawing.Size(269, 82)
        Me.光伏发电默认系数.TabIndex = 22
        Me.光伏发电默认系数.Text = "光伏默认系数"
        Me.光伏发电默认系数.UseVisualStyleBackColor = False
        '
        '蓄电池默认系数
        '
        Me.蓄电池默认系数.BackColor = System.Drawing.SystemColors.ControlLight
        Me.蓄电池默认系数.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.蓄电池默认系数.Location = New System.Drawing.Point(996, 650)
        Me.蓄电池默认系数.Margin = New System.Windows.Forms.Padding(6)
        Me.蓄电池默认系数.Name = "蓄电池默认系数"
        Me.蓄电池默认系数.Size = New System.Drawing.Size(269, 82)
        Me.蓄电池默认系数.TabIndex = 23
        Me.蓄电池默认系数.Text = "蓄电池默认系数"
        Me.蓄电池默认系数.UseVisualStyleBackColor = False
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label14.Location = New System.Drawing.Point(542, 22)
        Me.Label14.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(708, 50)
        Me.Label14.TabIndex = 32
        Me.Label14.Text = "光伏发电OR蓄电池计算开始年份和结束年份输入"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'gfjsnf
        '
        Me.gfjsnf.Location = New System.Drawing.Point(1123, 85)
        Me.gfjsnf.Margin = New System.Windows.Forms.Padding(6)
        Me.gfjsnf.Name = "gfjsnf"
        Me.gfjsnf.Size = New System.Drawing.Size(142, 35)
        Me.gfjsnf.TabIndex = 33
        Me.gfjsnf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'xdcjsnf
        '
        Me.xdcjsnf.Location = New System.Drawing.Point(1123, 147)
        Me.xdcjsnf.Margin = New System.Windows.Forms.Padding(6)
        Me.xdcjsnf.Name = "xdcjsnf"
        Me.xdcjsnf.Size = New System.Drawing.Size(142, 35)
        Me.xdcjsnf.TabIndex = 34
        Me.xdcjsnf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label15.Location = New System.Drawing.Point(917, 79)
        Me.Label15.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(194, 56)
        Me.Label15.TabIndex = 35
        Me.Label15.Text = "光伏结束年份"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label16.Location = New System.Drawing.Point(913, 141)
        Me.Label16.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(198, 56)
        Me.Label16.TabIndex = 36
        Me.Label16.Text = "蓄电池结束年份"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.Location = New System.Drawing.Point(539, 141)
        Me.Label17.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(203, 56)
        Me.Label17.TabIndex = 38
        Me.Label17.Text = "蓄电池开始年份"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label18.Location = New System.Drawing.Point(552, 79)
        Me.Label18.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(176, 56)
        Me.Label18.TabIndex = 37
        Me.Label18.Text = "光伏开始年份"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'xdcksnf
        '
        Me.xdcksnf.Location = New System.Drawing.Point(748, 147)
        Me.xdcksnf.Margin = New System.Windows.Forms.Padding(6)
        Me.xdcksnf.Name = "xdcksnf"
        Me.xdcksnf.Size = New System.Drawing.Size(142, 35)
        Me.xdcksnf.TabIndex = 40
        Me.xdcksnf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'gfksnf
        '
        Me.gfksnf.Location = New System.Drawing.Point(748, 85)
        Me.gfksnf.Margin = New System.Windows.Forms.Padding(6)
        Me.gfksnf.Name = "gfksnf"
        Me.gfksnf.Size = New System.Drawing.Size(142, 35)
        Me.gfksnf.TabIndex = 39
        Me.gfksnf.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'CheckBox1
        '
        Me.CheckBox1.AutoSize = True
        Me.CheckBox1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox1.Location = New System.Drawing.Point(584, 212)
        Me.CheckBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox1.Name = "CheckBox1"
        Me.CheckBox1.Size = New System.Drawing.Size(594, 28)
        Me.CheckBox1.TabIndex = 76
        Me.CheckBox1.Text = "光伏或者蓄电池逐年衰减率是否乘以投产月份比例？"
        Me.CheckBox1.UseVisualStyleBackColor = True
        '
        '蓄电池购电
        '
        Me.蓄电池购电.BackColor = System.Drawing.SystemColors.ControlLight
        Me.蓄电池购电.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.蓄电池购电.Location = New System.Drawing.Point(770, 650)
        Me.蓄电池购电.Margin = New System.Windows.Forms.Padding(6)
        Me.蓄电池购电.Name = "蓄电池购电"
        Me.蓄电池购电.Size = New System.Drawing.Size(206, 82)
        Me.蓄电池购电.TabIndex = 77
        Me.蓄电池购电.Text = "蓄电池购电"
        Me.蓄电池购电.UseVisualStyleBackColor = False
        '
        '清空输入
        '
        Me.清空输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空输入.Location = New System.Drawing.Point(277, 650)
        Me.清空输入.Margin = New System.Windows.Forms.Padding(6)
        Me.清空输入.Name = "清空输入"
        Me.清空输入.Size = New System.Drawing.Size(233, 82)
        Me.清空输入.TabIndex = 220
        Me.清空输入.Text = "清空输入"
        Me.清空输入.UseVisualStyleBackColor = False
        '
        '添加输入
        '
        Me.添加输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.添加输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.添加输入.Location = New System.Drawing.Point(28, 650)
        Me.添加输入.Margin = New System.Windows.Forms.Padding(6)
        Me.添加输入.Name = "添加输入"
        Me.添加输入.Size = New System.Drawing.Size(233, 82)
        Me.添加输入.TabIndex = 219
        Me.添加输入.Text = "添加输入"
        Me.添加输入.UseVisualStyleBackColor = False
        '
        '衰减率列表
        '
        Me.衰减率列表.FormattingEnabled = True
        Me.衰减率列表.ItemHeight = 24
        Me.衰减率列表.Location = New System.Drawing.Point(321, 141)
        Me.衰减率列表.Name = "衰减率列表"
        Me.衰减率列表.Size = New System.Drawing.Size(142, 436)
        Me.衰减率列表.TabIndex = 222
        '
        '开始年份列表
        '
        Me.开始年份列表.FormattingEnabled = True
        Me.开始年份列表.ItemHeight = 24
        Me.开始年份列表.Location = New System.Drawing.Point(72, 141)
        Me.开始年份列表.Name = "开始年份列表"
        Me.开始年份列表.Size = New System.Drawing.Size(142, 436)
        Me.开始年份列表.TabIndex = 221
        '
        '设置逐年衰减计算方式
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1300, 761)
        Me.Controls.Add(Me.衰减率列表)
        Me.Controls.Add(Me.开始年份列表)
        Me.Controls.Add(Me.清空输入)
        Me.Controls.Add(Me.添加输入)
        Me.Controls.Add(Me.蓄电池购电)
        Me.Controls.Add(Me.CheckBox1)
        Me.Controls.Add(Me.xdcksnf)
        Me.Controls.Add(Me.gfksnf)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.xdcjsnf)
        Me.Controls.Add(Me.gfjsnf)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.蓄电池默认系数)
        Me.Controls.Add(Me.光伏发电默认系数)
        Me.Controls.Add(Me.清空窗体)
        Me.Controls.Add(Me.蓄电池供电)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.光伏发电)
        Me.Controls.Add(Me.Label7)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.衰减率tmp)
        Me.Controls.Add(Me.开始年份tmp)
        Me.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Margin = New System.Windows.Forms.Padding(6)
        Me.Name = "设置逐年衰减计算方式"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置逐年衰减计算方式"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents 开始年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents 衰减率tmp As System.Windows.Forms.TextBox
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents 光伏发电 As System.Windows.Forms.Button
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents 蓄电池供电 As System.Windows.Forms.Button
    Friend WithEvents 清空窗体 As System.Windows.Forms.Button
    Friend WithEvents 光伏发电默认系数 As System.Windows.Forms.Button
    Friend WithEvents 蓄电池默认系数 As System.Windows.Forms.Button
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents gfjsnf As System.Windows.Forms.TextBox
    Friend WithEvents xdcjsnf As System.Windows.Forms.TextBox
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents xdcksnf As System.Windows.Forms.TextBox
    Friend WithEvents gfksnf As System.Windows.Forms.TextBox
    Friend WithEvents CheckBox1 As System.Windows.Forms.CheckBox
    Friend WithEvents 蓄电池购电 As System.Windows.Forms.Button
    Friend WithEvents 清空输入 As System.Windows.Forms.Button
    Friend WithEvents 添加输入 As System.Windows.Forms.Button
    Friend WithEvents 衰减率列表 As System.Windows.Forms.ListBox
    Friend WithEvents 开始年份列表 As System.Windows.Forms.ListBox
End Class
