<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 设置材料费和其它费计算方式
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
        Me.计算 = New System.Windows.Forms.Button()
        Me.风电 = New System.Windows.Forms.Button()
        Me.光伏 = New System.Windows.Forms.Button()
        Me.暖通 = New System.Windows.Forms.Button()
        Me.垃圾发电 = New System.Windows.Forms.Button()
        Me.燃机 = New System.Windows.Forms.Button()
        Me.燃煤 = New System.Windows.Forms.Button()
        Me.KCBL_LJFD = New System.Windows.Forms.TextBox()
        Me.Label36 = New System.Windows.Forms.Label()
        Me.YYNX_LJFD = New System.Windows.Forms.TextBox()
        Me.Label37 = New System.Windows.Forms.Label()
        Me.lajifadian = New System.Windows.Forms.CheckBox()
        Me.Label35 = New System.Windows.Forms.Label()
        Me.Label8 = New System.Windows.Forms.Label()
        Me.KCBL_FD = New System.Windows.Forms.TextBox()
        Me.Label33 = New System.Windows.Forms.Label()
        Me.YYNX_FD = New System.Windows.Forms.TextBox()
        Me.Label34 = New System.Windows.Forms.Label()
        Me.KCBL_GF = New System.Windows.Forms.TextBox()
        Me.Label31 = New System.Windows.Forms.Label()
        Me.YYNX_GF = New System.Windows.Forms.TextBox()
        Me.Label32 = New System.Windows.Forms.Label()
        Me.KCBL_NT = New System.Windows.Forms.TextBox()
        Me.Label29 = New System.Windows.Forms.Label()
        Me.YYNX_NT = New System.Windows.Forms.TextBox()
        Me.Label30 = New System.Windows.Forms.Label()
        Me.KCBL_RM = New System.Windows.Forms.TextBox()
        Me.Label24 = New System.Windows.Forms.Label()
        Me.YYNX_RM = New System.Windows.Forms.TextBox()
        Me.Label27 = New System.Windows.Forms.Label()
        Me.KCBL_RJ = New System.Windows.Forms.TextBox()
        Me.Label28 = New System.Windows.Forms.Label()
        Me.YYNX_RJ = New System.Windows.Forms.TextBox()
        Me.Label26 = New System.Windows.Forms.Label()
        Me.Label25 = New System.Windows.Forms.Label()
        Me.fengdian = New System.Windows.Forms.CheckBox()
        Me.guangfu = New System.Windows.Forms.CheckBox()
        Me.nuantong = New System.Windows.Forms.CheckBox()
        Me.ranmei = New System.Windows.Forms.CheckBox()
        Me.ranji = New System.Windows.Forms.CheckBox()
        Me.设置剔除 = New System.Windows.Forms.Button()
        Me.清空窗体 = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.结束费率tmp = New System.Windows.Forms.TextBox()
        Me.开始费率tmp = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.结束年份tmp = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.开始年份tmp = New System.Windows.Forms.TextBox()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.蓄电池 = New System.Windows.Forms.Button()
        Me.KCBL_XDC = New System.Windows.Forms.TextBox()
        Me.Label39 = New System.Windows.Forms.Label()
        Me.YYNX_XDC = New System.Windows.Forms.TextBox()
        Me.Label40 = New System.Windows.Forms.Label()
        Me.xudianchi = New System.Windows.Forms.CheckBox()
        Me.清空输入 = New System.Windows.Forms.Button()
        Me.添加输入 = New System.Windows.Forms.Button()
        Me.开始年份列表 = New System.Windows.Forms.ListBox()
        Me.结束年份列表 = New System.Windows.Forms.ListBox()
        Me.开始费率列表 = New System.Windows.Forms.ListBox()
        Me.结束费率列表 = New System.Windows.Forms.ListBox()
        Me.SuspendLayout()
        '
        '计算
        '
        Me.计算.BackColor = System.Drawing.SystemColors.ControlLight
        Me.计算.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.计算.Location = New System.Drawing.Point(1433, 522)
        Me.计算.Margin = New System.Windows.Forms.Padding(6)
        Me.计算.Name = "计算"
        Me.计算.Size = New System.Drawing.Size(224, 82)
        Me.计算.TabIndex = 198
        Me.计算.Text = "计算"
        Me.计算.UseVisualStyleBackColor = False
        '
        '风电
        '
        Me.风电.BackColor = System.Drawing.SystemColors.ControlLight
        Me.风电.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.风电.Location = New System.Drawing.Point(233, 812)
        Me.风电.Margin = New System.Windows.Forms.Padding(6)
        Me.风电.Name = "风电"
        Me.风电.Size = New System.Drawing.Size(182, 82)
        Me.风电.TabIndex = 195
        Me.风电.Text = "风电"
        Me.风电.UseVisualStyleBackColor = False
        '
        '光伏
        '
        Me.光伏.BackColor = System.Drawing.SystemColors.ControlLight
        Me.光伏.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.光伏.Location = New System.Drawing.Point(39, 812)
        Me.光伏.Margin = New System.Windows.Forms.Padding(6)
        Me.光伏.Name = "光伏"
        Me.光伏.Size = New System.Drawing.Size(182, 82)
        Me.光伏.TabIndex = 194
        Me.光伏.Text = "光伏"
        Me.光伏.UseVisualStyleBackColor = False
        '
        '暖通
        '
        Me.暖通.BackColor = System.Drawing.SystemColors.ControlLight
        Me.暖通.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.暖通.Location = New System.Drawing.Point(427, 716)
        Me.暖通.Margin = New System.Windows.Forms.Padding(6)
        Me.暖通.Name = "暖通"
        Me.暖通.Size = New System.Drawing.Size(182, 82)
        Me.暖通.TabIndex = 193
        Me.暖通.Text = "暖通"
        Me.暖通.UseVisualStyleBackColor = False
        '
        '垃圾发电
        '
        Me.垃圾发电.BackColor = System.Drawing.SystemColors.ControlLight
        Me.垃圾发电.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.垃圾发电.Location = New System.Drawing.Point(621, 716)
        Me.垃圾发电.Margin = New System.Windows.Forms.Padding(6)
        Me.垃圾发电.Name = "垃圾发电"
        Me.垃圾发电.Size = New System.Drawing.Size(182, 82)
        Me.垃圾发电.TabIndex = 192
        Me.垃圾发电.Text = "垃圾发电"
        Me.垃圾发电.UseVisualStyleBackColor = False
        '
        '燃机
        '
        Me.燃机.BackColor = System.Drawing.SystemColors.ControlLight
        Me.燃机.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.燃机.Location = New System.Drawing.Point(39, 716)
        Me.燃机.Margin = New System.Windows.Forms.Padding(6)
        Me.燃机.Name = "燃机"
        Me.燃机.Size = New System.Drawing.Size(182, 82)
        Me.燃机.TabIndex = 191
        Me.燃机.Text = "燃机"
        Me.燃机.UseVisualStyleBackColor = False
        '
        '燃煤
        '
        Me.燃煤.BackColor = System.Drawing.SystemColors.ControlLight
        Me.燃煤.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.燃煤.Location = New System.Drawing.Point(233, 716)
        Me.燃煤.Margin = New System.Windows.Forms.Padding(6)
        Me.燃煤.Name = "燃煤"
        Me.燃煤.Size = New System.Drawing.Size(182, 82)
        Me.燃煤.TabIndex = 190
        Me.燃煤.Text = "燃煤"
        Me.燃煤.UseVisualStyleBackColor = False
        '
        'KCBL_LJFD
        '
        Me.KCBL_LJFD.Location = New System.Drawing.Point(1519, 459)
        Me.KCBL_LJFD.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_LJFD.Name = "KCBL_LJFD"
        Me.KCBL_LJFD.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_LJFD.TabIndex = 189
        Me.KCBL_LJFD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label36
        '
        Me.Label36.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label36.Location = New System.Drawing.Point(1353, 449)
        Me.Label36.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(154, 56)
        Me.Label36.TabIndex = 188
        Me.Label36.Text = "扣除比例(%)"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_LJFD
        '
        Me.YYNX_LJFD.Location = New System.Drawing.Point(1190, 459)
        Me.YYNX_LJFD.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_LJFD.Name = "YYNX_LJFD"
        Me.YYNX_LJFD.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_LJFD.TabIndex = 187
        Me.YYNX_LJFD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label37
        '
        Me.Label37.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label37.Location = New System.Drawing.Point(1029, 450)
        Me.Label37.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label37.Name = "Label37"
        Me.Label37.Size = New System.Drawing.Size(158, 56)
        Me.Label37.TabIndex = 186
        Me.Label37.Text = "运营年限(年)"
        Me.Label37.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'lajifadian
        '
        Me.lajifadian.AutoSize = True
        Me.lajifadian.Location = New System.Drawing.Point(881, 465)
        Me.lajifadian.Name = "lajifadian"
        Me.lajifadian.Size = New System.Drawing.Size(138, 28)
        Me.lajifadian.TabIndex = 185
        Me.lajifadian.Text = "垃圾发电"
        Me.lajifadian.UseVisualStyleBackColor = True
        '
        'Label35
        '
        Me.Label35.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label35.Location = New System.Drawing.Point(1011, 25)
        Me.Label35.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label35.Name = "Label35"
        Me.Label35.Size = New System.Drawing.Size(528, 50)
        Me.Label35.TabIndex = 184
        Me.Label35.Text = "设备材料费率OR其它费率计算特殊设置"
        Me.Label35.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label8
        '
        Me.Label8.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label8.Location = New System.Drawing.Point(189, 25)
        Me.Label8.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(528, 50)
        Me.Label8.TabIndex = 183
        Me.Label8.Text = "设备逐年材料费率OR其它费率计算特殊设置"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_FD
        '
        Me.KCBL_FD.Location = New System.Drawing.Point(1519, 355)
        Me.KCBL_FD.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_FD.Name = "KCBL_FD"
        Me.KCBL_FD.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_FD.TabIndex = 182
        Me.KCBL_FD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label33.Location = New System.Drawing.Point(1353, 345)
        Me.Label33.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(154, 56)
        Me.Label33.TabIndex = 181
        Me.Label33.Text = "扣除比例(%)"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_FD
        '
        Me.YYNX_FD.Location = New System.Drawing.Point(1190, 355)
        Me.YYNX_FD.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_FD.Name = "YYNX_FD"
        Me.YYNX_FD.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_FD.TabIndex = 180
        Me.YYNX_FD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label34
        '
        Me.Label34.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label34.Location = New System.Drawing.Point(1029, 346)
        Me.Label34.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(158, 56)
        Me.Label34.TabIndex = 179
        Me.Label34.Text = "运营年限(年)"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_GF
        '
        Me.KCBL_GF.Location = New System.Drawing.Point(1519, 303)
        Me.KCBL_GF.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_GF.Name = "KCBL_GF"
        Me.KCBL_GF.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_GF.TabIndex = 178
        Me.KCBL_GF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label31.Location = New System.Drawing.Point(1353, 293)
        Me.Label31.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(154, 56)
        Me.Label31.TabIndex = 177
        Me.Label31.Text = "扣除比例(%)"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_GF
        '
        Me.YYNX_GF.Location = New System.Drawing.Point(1190, 303)
        Me.YYNX_GF.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_GF.Name = "YYNX_GF"
        Me.YYNX_GF.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_GF.TabIndex = 176
        Me.YYNX_GF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label32.Location = New System.Drawing.Point(1029, 294)
        Me.Label32.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(158, 56)
        Me.Label32.TabIndex = 175
        Me.Label32.Text = "运营年限(年)"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_NT
        '
        Me.KCBL_NT.Location = New System.Drawing.Point(1519, 247)
        Me.KCBL_NT.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_NT.Name = "KCBL_NT"
        Me.KCBL_NT.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_NT.TabIndex = 174
        Me.KCBL_NT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label29.Location = New System.Drawing.Point(1353, 237)
        Me.Label29.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(154, 56)
        Me.Label29.TabIndex = 173
        Me.Label29.Text = "扣除比例(%)"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_NT
        '
        Me.YYNX_NT.Location = New System.Drawing.Point(1190, 247)
        Me.YYNX_NT.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_NT.Name = "YYNX_NT"
        Me.YYNX_NT.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_NT.TabIndex = 172
        Me.YYNX_NT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label30
        '
        Me.Label30.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label30.Location = New System.Drawing.Point(1029, 238)
        Me.Label30.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(158, 56)
        Me.Label30.TabIndex = 171
        Me.Label30.Text = "运营年限(年)"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_RM
        '
        Me.KCBL_RM.Location = New System.Drawing.Point(1519, 195)
        Me.KCBL_RM.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_RM.Name = "KCBL_RM"
        Me.KCBL_RM.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_RM.TabIndex = 170
        Me.KCBL_RM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label24.Location = New System.Drawing.Point(1353, 185)
        Me.Label24.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(154, 56)
        Me.Label24.TabIndex = 169
        Me.Label24.Text = "扣除比例(%)"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_RM
        '
        Me.YYNX_RM.Location = New System.Drawing.Point(1190, 195)
        Me.YYNX_RM.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_RM.Name = "YYNX_RM"
        Me.YYNX_RM.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_RM.TabIndex = 168
        Me.YYNX_RM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label27.Location = New System.Drawing.Point(1029, 186)
        Me.Label27.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(158, 56)
        Me.Label27.TabIndex = 167
        Me.Label27.Text = "运营年限(年)"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_RJ
        '
        Me.KCBL_RJ.Location = New System.Drawing.Point(1519, 139)
        Me.KCBL_RJ.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_RJ.Name = "KCBL_RJ"
        Me.KCBL_RJ.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_RJ.TabIndex = 166
        Me.KCBL_RJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label28.Location = New System.Drawing.Point(1353, 129)
        Me.Label28.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(154, 56)
        Me.Label28.TabIndex = 165
        Me.Label28.Text = "扣除比例(%)"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_RJ
        '
        Me.YYNX_RJ.Location = New System.Drawing.Point(1190, 139)
        Me.YYNX_RJ.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_RJ.Name = "YYNX_RJ"
        Me.YYNX_RJ.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_RJ.TabIndex = 164
        Me.YYNX_RJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label26.Location = New System.Drawing.Point(1029, 130)
        Me.Label26.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label26.Name = "Label26"
        Me.Label26.Size = New System.Drawing.Size(158, 56)
        Me.Label26.TabIndex = 163
        Me.Label26.Text = "运营年限(年)"
        Me.Label26.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label25
        '
        Me.Label25.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label25.Location = New System.Drawing.Point(906, 75)
        Me.Label25.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label25.Name = "Label25"
        Me.Label25.Size = New System.Drawing.Size(731, 50)
        Me.Label25.TabIndex = 162
        Me.Label25.Text = "已经报废的设备材料费OR其它费计算基数剔除"
        Me.Label25.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'fengdian
        '
        Me.fengdian.AutoSize = True
        Me.fengdian.Location = New System.Drawing.Point(881, 361)
        Me.fengdian.Name = "fengdian"
        Me.fengdian.Size = New System.Drawing.Size(90, 28)
        Me.fengdian.TabIndex = 161
        Me.fengdian.Text = "风电"
        Me.fengdian.UseVisualStyleBackColor = True
        '
        'guangfu
        '
        Me.guangfu.AutoSize = True
        Me.guangfu.Location = New System.Drawing.Point(881, 309)
        Me.guangfu.Name = "guangfu"
        Me.guangfu.Size = New System.Drawing.Size(90, 28)
        Me.guangfu.TabIndex = 160
        Me.guangfu.Text = "光伏"
        Me.guangfu.UseVisualStyleBackColor = True
        '
        'nuantong
        '
        Me.nuantong.AutoSize = True
        Me.nuantong.Location = New System.Drawing.Point(881, 253)
        Me.nuantong.Name = "nuantong"
        Me.nuantong.Size = New System.Drawing.Size(90, 28)
        Me.nuantong.TabIndex = 159
        Me.nuantong.Text = "暖通"
        Me.nuantong.UseVisualStyleBackColor = True
        '
        'ranmei
        '
        Me.ranmei.AutoSize = True
        Me.ranmei.Location = New System.Drawing.Point(881, 199)
        Me.ranmei.Name = "ranmei"
        Me.ranmei.Size = New System.Drawing.Size(90, 28)
        Me.ranmei.TabIndex = 158
        Me.ranmei.Text = "燃煤"
        Me.ranmei.UseVisualStyleBackColor = True
        '
        'ranji
        '
        Me.ranji.AutoSize = True
        Me.ranji.Location = New System.Drawing.Point(881, 145)
        Me.ranji.Name = "ranji"
        Me.ranji.Size = New System.Drawing.Size(90, 28)
        Me.ranji.TabIndex = 157
        Me.ranji.Text = "燃机"
        Me.ranji.UseVisualStyleBackColor = True
        '
        '设置剔除
        '
        Me.设置剔除.BackColor = System.Drawing.SystemColors.ControlLight
        Me.设置剔除.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.设置剔除.Location = New System.Drawing.Point(877, 522)
        Me.设置剔除.Margin = New System.Windows.Forms.Padding(6)
        Me.设置剔除.Name = "设置剔除"
        Me.设置剔除.Size = New System.Drawing.Size(224, 82)
        Me.设置剔除.TabIndex = 148
        Me.设置剔除.Text = "设置剔除"
        Me.设置剔除.UseVisualStyleBackColor = False
        '
        '清空窗体
        '
        Me.清空窗体.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空窗体.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空窗体.Location = New System.Drawing.Point(1153, 522)
        Me.清空窗体.Margin = New System.Windows.Forms.Padding(6)
        Me.清空窗体.Name = "清空窗体"
        Me.清空窗体.Size = New System.Drawing.Size(224, 82)
        Me.清空窗体.TabIndex = 147
        Me.清空窗体.Text = "清空窗体"
        Me.清空窗体.UseVisualStyleBackColor = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(881, 673)
        Me.RichTextBox1.Margin = New System.Windows.Forms.Padding(6)
        Me.RichTextBox1.Name = "RichTextBox1"
        Me.RichTextBox1.ReadOnly = True
        Me.RichTextBox1.Size = New System.Drawing.Size(780, 227)
        Me.RichTextBox1.TabIndex = 146
        Me.RichTextBox1.Text = ""
        '
        'Label19
        '
        Me.Label19.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label19.Location = New System.Drawing.Point(980, 617)
        Me.Label19.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(583, 50)
        Me.Label19.TabIndex = 145
        Me.Label19.Text = "逐年材料费率OR其它费率计算结果显示"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '结束费率tmp
        '
        Me.结束费率tmp.Location = New System.Drawing.Point(643, 525)
        Me.结束费率tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.结束费率tmp.Name = "结束费率tmp"
        Me.结束费率tmp.Size = New System.Drawing.Size(142, 35)
        Me.结束费率tmp.TabIndex = 120
        Me.结束费率tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '开始费率tmp
        '
        Me.开始费率tmp.Location = New System.Drawing.Point(419, 525)
        Me.开始费率tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.开始费率tmp.Name = "开始费率tmp"
        Me.开始费率tmp.Size = New System.Drawing.Size(142, 35)
        Me.开始费率tmp.TabIndex = 119
        Me.开始费率tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(622, 77)
        Me.Label2.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(196, 56)
        Me.Label2.TabIndex = 118
        Me.Label2.Text = "结束年份费率"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label1.Location = New System.Drawing.Point(401, 77)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 56)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "开始年份费率"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '结束年份tmp
        '
        Me.结束年份tmp.Location = New System.Drawing.Point(213, 525)
        Me.结束年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.结束年份tmp.Name = "结束年份tmp"
        Me.结束年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.结束年份tmp.TabIndex = 116
        Me.结束年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(225, 73)
        Me.Label4.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(132, 56)
        Me.Label4.TabIndex = 115
        Me.Label4.Text = "结束年份"
        Me.Label4.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label3
        '
        Me.Label3.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label3.Location = New System.Drawing.Point(51, 76)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(134, 56)
        Me.Label3.TabIndex = 114
        Me.Label3.Text = "开始年份"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '开始年份tmp
        '
        Me.开始年份tmp.Location = New System.Drawing.Point(39, 525)
        Me.开始年份tmp.Margin = New System.Windows.Forms.Padding(6)
        Me.开始年份tmp.Name = "开始年份tmp"
        Me.开始年份tmp.Size = New System.Drawing.Size(142, 35)
        Me.开始年份tmp.TabIndex = 113
        Me.开始年份tmp.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"材料费率", "其它费率"})
        Me.ComboBox1.Location = New System.Drawing.Point(542, 667)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(243, 32)
        Me.ComboBox1.TabIndex = 208
        '
        'Label38
        '
        Me.Label38.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label38.Location = New System.Drawing.Point(35, 657)
        Me.Label38.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(445, 50)
        Me.Label38.TabIndex = 207
        Me.Label38.Text = "选择需要设置的费率类型"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '蓄电池
        '
        Me.蓄电池.BackColor = System.Drawing.SystemColors.ControlLight
        Me.蓄电池.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.蓄电池.Location = New System.Drawing.Point(427, 812)
        Me.蓄电池.Margin = New System.Windows.Forms.Padding(6)
        Me.蓄电池.Name = "蓄电池"
        Me.蓄电池.Size = New System.Drawing.Size(182, 82)
        Me.蓄电池.TabIndex = 209
        Me.蓄电池.Text = "蓄电池"
        Me.蓄电池.UseVisualStyleBackColor = False
        '
        'KCBL_XDC
        '
        Me.KCBL_XDC.Location = New System.Drawing.Point(1519, 406)
        Me.KCBL_XDC.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_XDC.Name = "KCBL_XDC"
        Me.KCBL_XDC.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_XDC.TabIndex = 214
        Me.KCBL_XDC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label39
        '
        Me.Label39.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label39.Location = New System.Drawing.Point(1353, 396)
        Me.Label39.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label39.Name = "Label39"
        Me.Label39.Size = New System.Drawing.Size(154, 56)
        Me.Label39.TabIndex = 213
        Me.Label39.Text = "扣除比例(%)"
        Me.Label39.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_XDC
        '
        Me.YYNX_XDC.Location = New System.Drawing.Point(1190, 406)
        Me.YYNX_XDC.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_XDC.Name = "YYNX_XDC"
        Me.YYNX_XDC.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_XDC.TabIndex = 212
        Me.YYNX_XDC.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label40
        '
        Me.Label40.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label40.Location = New System.Drawing.Point(1029, 397)
        Me.Label40.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label40.Name = "Label40"
        Me.Label40.Size = New System.Drawing.Size(158, 56)
        Me.Label40.TabIndex = 211
        Me.Label40.Text = "运营年限(年)"
        Me.Label40.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'xudianchi
        '
        Me.xudianchi.AutoSize = True
        Me.xudianchi.Location = New System.Drawing.Point(881, 412)
        Me.xudianchi.Name = "xudianchi"
        Me.xudianchi.Size = New System.Drawing.Size(114, 28)
        Me.xudianchi.TabIndex = 210
        Me.xudianchi.Text = "蓄电池"
        Me.xudianchi.UseVisualStyleBackColor = True
        '
        '清空输入
        '
        Me.清空输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.清空输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.清空输入.Location = New System.Drawing.Point(488, 577)
        Me.清空输入.Margin = New System.Windows.Forms.Padding(6)
        Me.清空输入.Name = "清空输入"
        Me.清空输入.Size = New System.Drawing.Size(233, 72)
        Me.清空输入.TabIndex = 216
        Me.清空输入.Text = "清空输入"
        Me.清空输入.UseVisualStyleBackColor = False
        '
        '添加输入
        '
        Me.添加输入.BackColor = System.Drawing.SystemColors.ControlLight
        Me.添加输入.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.添加输入.Location = New System.Drawing.Point(86, 576)
        Me.添加输入.Margin = New System.Windows.Forms.Padding(6)
        Me.添加输入.Name = "添加输入"
        Me.添加输入.Size = New System.Drawing.Size(233, 72)
        Me.添加输入.TabIndex = 215
        Me.添加输入.Text = "添加输入"
        Me.添加输入.UseVisualStyleBackColor = False
        '
        '开始年份列表
        '
        Me.开始年份列表.FormattingEnabled = True
        Me.开始年份列表.ItemHeight = 24
        Me.开始年份列表.Location = New System.Drawing.Point(39, 139)
        Me.开始年份列表.Name = "开始年份列表"
        Me.开始年份列表.Size = New System.Drawing.Size(142, 364)
        Me.开始年份列表.TabIndex = 217
        '
        '结束年份列表
        '
        Me.结束年份列表.FormattingEnabled = True
        Me.结束年份列表.ItemHeight = 24
        Me.结束年份列表.Location = New System.Drawing.Point(213, 139)
        Me.结束年份列表.Name = "结束年份列表"
        Me.结束年份列表.Size = New System.Drawing.Size(142, 364)
        Me.结束年份列表.TabIndex = 218
        '
        '开始费率列表
        '
        Me.开始费率列表.FormattingEnabled = True
        Me.开始费率列表.ItemHeight = 24
        Me.开始费率列表.Location = New System.Drawing.Point(419, 139)
        Me.开始费率列表.Name = "开始费率列表"
        Me.开始费率列表.Size = New System.Drawing.Size(142, 364)
        Me.开始费率列表.TabIndex = 219
        '
        '结束费率列表
        '
        Me.结束费率列表.FormattingEnabled = True
        Me.结束费率列表.ItemHeight = 24
        Me.结束费率列表.Location = New System.Drawing.Point(643, 139)
        Me.结束费率列表.Name = "结束费率列表"
        Me.结束费率列表.Size = New System.Drawing.Size(142, 364)
        Me.结束费率列表.TabIndex = 220
        '
        '设置材料费和其它费计算方式
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1698, 930)
        Me.Controls.Add(Me.结束费率列表)
        Me.Controls.Add(Me.开始费率列表)
        Me.Controls.Add(Me.结束年份列表)
        Me.Controls.Add(Me.开始年份列表)
        Me.Controls.Add(Me.清空输入)
        Me.Controls.Add(Me.添加输入)
        Me.Controls.Add(Me.KCBL_XDC)
        Me.Controls.Add(Me.Label39)
        Me.Controls.Add(Me.YYNX_XDC)
        Me.Controls.Add(Me.Label40)
        Me.Controls.Add(Me.xudianchi)
        Me.Controls.Add(Me.蓄电池)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.计算)
        Me.Controls.Add(Me.风电)
        Me.Controls.Add(Me.光伏)
        Me.Controls.Add(Me.暖通)
        Me.Controls.Add(Me.垃圾发电)
        Me.Controls.Add(Me.燃机)
        Me.Controls.Add(Me.燃煤)
        Me.Controls.Add(Me.KCBL_LJFD)
        Me.Controls.Add(Me.Label36)
        Me.Controls.Add(Me.YYNX_LJFD)
        Me.Controls.Add(Me.Label37)
        Me.Controls.Add(Me.lajifadian)
        Me.Controls.Add(Me.Label35)
        Me.Controls.Add(Me.Label8)
        Me.Controls.Add(Me.KCBL_FD)
        Me.Controls.Add(Me.Label33)
        Me.Controls.Add(Me.YYNX_FD)
        Me.Controls.Add(Me.Label34)
        Me.Controls.Add(Me.KCBL_GF)
        Me.Controls.Add(Me.Label31)
        Me.Controls.Add(Me.YYNX_GF)
        Me.Controls.Add(Me.Label32)
        Me.Controls.Add(Me.KCBL_NT)
        Me.Controls.Add(Me.Label29)
        Me.Controls.Add(Me.YYNX_NT)
        Me.Controls.Add(Me.Label30)
        Me.Controls.Add(Me.KCBL_RM)
        Me.Controls.Add(Me.Label24)
        Me.Controls.Add(Me.YYNX_RM)
        Me.Controls.Add(Me.Label27)
        Me.Controls.Add(Me.KCBL_RJ)
        Me.Controls.Add(Me.Label28)
        Me.Controls.Add(Me.YYNX_RJ)
        Me.Controls.Add(Me.Label26)
        Me.Controls.Add(Me.Label25)
        Me.Controls.Add(Me.fengdian)
        Me.Controls.Add(Me.guangfu)
        Me.Controls.Add(Me.nuantong)
        Me.Controls.Add(Me.ranmei)
        Me.Controls.Add(Me.ranji)
        Me.Controls.Add(Me.设置剔除)
        Me.Controls.Add(Me.清空窗体)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.结束费率tmp)
        Me.Controls.Add(Me.开始费率tmp)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.结束年份tmp)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.开始年份tmp)
        Me.Name = "设置材料费和其它费计算方式"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "设置材料费和其它费计算方式"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents 计算 As System.Windows.Forms.Button
    Friend WithEvents 风电 As System.Windows.Forms.Button
    Friend WithEvents 光伏 As System.Windows.Forms.Button
    Friend WithEvents 暖通 As System.Windows.Forms.Button
    Friend WithEvents 垃圾发电 As System.Windows.Forms.Button
    Friend WithEvents 燃机 As System.Windows.Forms.Button
    Friend WithEvents 燃煤 As System.Windows.Forms.Button
    Friend WithEvents KCBL_LJFD As System.Windows.Forms.TextBox
    Friend WithEvents Label36 As System.Windows.Forms.Label
    Friend WithEvents YYNX_LJFD As System.Windows.Forms.TextBox
    Friend WithEvents Label37 As System.Windows.Forms.Label
    Friend WithEvents lajifadian As System.Windows.Forms.CheckBox
    Friend WithEvents Label35 As System.Windows.Forms.Label
    Friend WithEvents Label8 As System.Windows.Forms.Label
    Friend WithEvents KCBL_FD As System.Windows.Forms.TextBox
    Friend WithEvents Label33 As System.Windows.Forms.Label
    Friend WithEvents YYNX_FD As System.Windows.Forms.TextBox
    Friend WithEvents Label34 As System.Windows.Forms.Label
    Friend WithEvents KCBL_GF As System.Windows.Forms.TextBox
    Friend WithEvents Label31 As System.Windows.Forms.Label
    Friend WithEvents YYNX_GF As System.Windows.Forms.TextBox
    Friend WithEvents Label32 As System.Windows.Forms.Label
    Friend WithEvents KCBL_NT As System.Windows.Forms.TextBox
    Friend WithEvents Label29 As System.Windows.Forms.Label
    Friend WithEvents YYNX_NT As System.Windows.Forms.TextBox
    Friend WithEvents Label30 As System.Windows.Forms.Label
    Friend WithEvents KCBL_RM As System.Windows.Forms.TextBox
    Friend WithEvents Label24 As System.Windows.Forms.Label
    Friend WithEvents YYNX_RM As System.Windows.Forms.TextBox
    Friend WithEvents Label27 As System.Windows.Forms.Label
    Friend WithEvents KCBL_RJ As System.Windows.Forms.TextBox
    Friend WithEvents Label28 As System.Windows.Forms.Label
    Friend WithEvents YYNX_RJ As System.Windows.Forms.TextBox
    Friend WithEvents Label26 As System.Windows.Forms.Label
    Friend WithEvents Label25 As System.Windows.Forms.Label
    Friend WithEvents fengdian As System.Windows.Forms.CheckBox
    Friend WithEvents guangfu As System.Windows.Forms.CheckBox
    Friend WithEvents nuantong As System.Windows.Forms.CheckBox
    Friend WithEvents ranmei As System.Windows.Forms.CheckBox
    Friend WithEvents ranji As System.Windows.Forms.CheckBox
    Friend WithEvents 设置剔除 As System.Windows.Forms.Button
    Friend WithEvents 清空窗体 As System.Windows.Forms.Button
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents 结束费率tmp As System.Windows.Forms.TextBox
    Friend WithEvents 开始费率tmp As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents 结束年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents 开始年份tmp As System.Windows.Forms.TextBox
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
    Friend WithEvents 蓄电池 As System.Windows.Forms.Button
    Friend WithEvents KCBL_XDC As System.Windows.Forms.TextBox
    Friend WithEvents Label39 As System.Windows.Forms.Label
    Friend WithEvents YYNX_XDC As System.Windows.Forms.TextBox
    Friend WithEvents Label40 As System.Windows.Forms.Label
    Friend WithEvents xudianchi As System.Windows.Forms.CheckBox
    Friend WithEvents 清空输入 As System.Windows.Forms.Button
    Friend WithEvents 添加输入 As System.Windows.Forms.Button
    Friend WithEvents 开始年份列表 As System.Windows.Forms.ListBox
    Friend WithEvents 结束年份列表 As System.Windows.Forms.ListBox
    Friend WithEvents 开始费率列表 As System.Windows.Forms.ListBox
    Friend WithEvents 结束费率列表 As System.Windows.Forms.ListBox
End Class
