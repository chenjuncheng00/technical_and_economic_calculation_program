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
        Me.jsfl5 = New System.Windows.Forms.TextBox()
        Me.ksfl5 = New System.Windows.Forms.TextBox()
        Me.jsnf5 = New System.Windows.Forms.TextBox()
        Me.Label22 = New System.Windows.Forms.Label()
        Me.Label23 = New System.Windows.Forms.Label()
        Me.ksnf5 = New System.Windows.Forms.TextBox()
        Me.设置剔除 = New System.Windows.Forms.Button()
        Me.清空窗体 = New System.Windows.Forms.Button()
        Me.RichTextBox1 = New System.Windows.Forms.RichTextBox()
        Me.Label19 = New System.Windows.Forms.Label()
        Me.jsfl4 = New System.Windows.Forms.TextBox()
        Me.ksfl4 = New System.Windows.Forms.TextBox()
        Me.jsnf4 = New System.Windows.Forms.TextBox()
        Me.Label17 = New System.Windows.Forms.Label()
        Me.Label18 = New System.Windows.Forms.Label()
        Me.ksnf4 = New System.Windows.Forms.TextBox()
        Me.jsfl3 = New System.Windows.Forms.TextBox()
        Me.ksfl3 = New System.Windows.Forms.TextBox()
        Me.jsnf3 = New System.Windows.Forms.TextBox()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.Label14 = New System.Windows.Forms.Label()
        Me.ksnf3 = New System.Windows.Forms.TextBox()
        Me.jsfl2 = New System.Windows.Forms.TextBox()
        Me.ksfl2 = New System.Windows.Forms.TextBox()
        Me.jsnf2 = New System.Windows.Forms.TextBox()
        Me.Label9 = New System.Windows.Forms.Label()
        Me.Label10 = New System.Windows.Forms.Label()
        Me.ksnf2 = New System.Windows.Forms.TextBox()
        Me.jsfl1 = New System.Windows.Forms.TextBox()
        Me.ksfl1 = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.jsnf1 = New System.Windows.Forms.TextBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.ksnf1 = New System.Windows.Forms.TextBox()
        Me.Label7 = New System.Windows.Forms.Label()
        Me.Label5 = New System.Windows.Forms.Label()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Label15 = New System.Windows.Forms.Label()
        Me.Label16 = New System.Windows.Forms.Label()
        Me.Label20 = New System.Windows.Forms.Label()
        Me.Label21 = New System.Windows.Forms.Label()
        Me.ComboBox1 = New System.Windows.Forms.ComboBox()
        Me.Label38 = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        '计算
        '
        Me.计算.BackColor = System.Drawing.SystemColors.ControlLight
        Me.计算.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.计算.Location = New System.Drawing.Point(1437, 492)
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
        Me.风电.Location = New System.Drawing.Point(323, 812)
        Me.风电.Margin = New System.Windows.Forms.Padding(6)
        Me.风电.Name = "风电"
        Me.风电.Size = New System.Drawing.Size(224, 82)
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
        Me.光伏.Size = New System.Drawing.Size(224, 82)
        Me.光伏.TabIndex = 194
        Me.光伏.Text = "光伏"
        Me.光伏.UseVisualStyleBackColor = False
        '
        '暖通
        '
        Me.暖通.BackColor = System.Drawing.SystemColors.ControlLight
        Me.暖通.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.暖通.Location = New System.Drawing.Point(594, 701)
        Me.暖通.Margin = New System.Windows.Forms.Padding(6)
        Me.暖通.Name = "暖通"
        Me.暖通.Size = New System.Drawing.Size(224, 82)
        Me.暖通.TabIndex = 193
        Me.暖通.Text = "暖通"
        Me.暖通.UseVisualStyleBackColor = False
        '
        '垃圾发电
        '
        Me.垃圾发电.BackColor = System.Drawing.SystemColors.ControlLight
        Me.垃圾发电.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.垃圾发电.Location = New System.Drawing.Point(594, 812)
        Me.垃圾发电.Margin = New System.Windows.Forms.Padding(6)
        Me.垃圾发电.Name = "垃圾发电"
        Me.垃圾发电.Size = New System.Drawing.Size(224, 82)
        Me.垃圾发电.TabIndex = 192
        Me.垃圾发电.Text = "垃圾发电"
        Me.垃圾发电.UseVisualStyleBackColor = False
        '
        '燃机
        '
        Me.燃机.BackColor = System.Drawing.SystemColors.ControlLight
        Me.燃机.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.燃机.Location = New System.Drawing.Point(39, 701)
        Me.燃机.Margin = New System.Windows.Forms.Padding(6)
        Me.燃机.Name = "燃机"
        Me.燃机.Size = New System.Drawing.Size(224, 82)
        Me.燃机.TabIndex = 191
        Me.燃机.Text = "燃机"
        Me.燃机.UseVisualStyleBackColor = False
        '
        '燃煤
        '
        Me.燃煤.BackColor = System.Drawing.SystemColors.ControlLight
        Me.燃煤.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.燃煤.Location = New System.Drawing.Point(323, 701)
        Me.燃煤.Margin = New System.Windows.Forms.Padding(6)
        Me.燃煤.Name = "燃煤"
        Me.燃煤.Size = New System.Drawing.Size(224, 82)
        Me.燃煤.TabIndex = 190
        Me.燃煤.Text = "燃煤"
        Me.燃煤.UseVisualStyleBackColor = False
        '
        'KCBL_LJFD
        '
        Me.KCBL_LJFD.Location = New System.Drawing.Point(1519, 419)
        Me.KCBL_LJFD.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_LJFD.Name = "KCBL_LJFD"
        Me.KCBL_LJFD.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_LJFD.TabIndex = 189
        Me.KCBL_LJFD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label36
        '
        Me.Label36.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label36.Location = New System.Drawing.Point(1353, 409)
        Me.Label36.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label36.Name = "Label36"
        Me.Label36.Size = New System.Drawing.Size(154, 56)
        Me.Label36.TabIndex = 188
        Me.Label36.Text = "扣除比例(%)"
        Me.Label36.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_LJFD
        '
        Me.YYNX_LJFD.Location = New System.Drawing.Point(1190, 419)
        Me.YYNX_LJFD.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_LJFD.Name = "YYNX_LJFD"
        Me.YYNX_LJFD.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_LJFD.TabIndex = 187
        Me.YYNX_LJFD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label37
        '
        Me.Label37.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label37.Location = New System.Drawing.Point(1029, 410)
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
        Me.lajifadian.Location = New System.Drawing.Point(881, 425)
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
        Me.Label8.Location = New System.Drawing.Point(209, 25)
        Me.Label8.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label8.Name = "Label8"
        Me.Label8.Size = New System.Drawing.Size(528, 50)
        Me.Label8.TabIndex = 183
        Me.Label8.Text = "设备逐年材料费率OR其它费率计算特殊设置"
        Me.Label8.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_FD
        '
        Me.KCBL_FD.Location = New System.Drawing.Point(1519, 369)
        Me.KCBL_FD.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_FD.Name = "KCBL_FD"
        Me.KCBL_FD.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_FD.TabIndex = 182
        Me.KCBL_FD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label33
        '
        Me.Label33.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label33.Location = New System.Drawing.Point(1353, 359)
        Me.Label33.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label33.Name = "Label33"
        Me.Label33.Size = New System.Drawing.Size(154, 56)
        Me.Label33.TabIndex = 181
        Me.Label33.Text = "扣除比例(%)"
        Me.Label33.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_FD
        '
        Me.YYNX_FD.Location = New System.Drawing.Point(1190, 369)
        Me.YYNX_FD.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_FD.Name = "YYNX_FD"
        Me.YYNX_FD.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_FD.TabIndex = 180
        Me.YYNX_FD.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label34
        '
        Me.Label34.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label34.Location = New System.Drawing.Point(1029, 360)
        Me.Label34.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label34.Name = "Label34"
        Me.Label34.Size = New System.Drawing.Size(158, 56)
        Me.Label34.TabIndex = 179
        Me.Label34.Text = "运营年限(年)"
        Me.Label34.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_GF
        '
        Me.KCBL_GF.Location = New System.Drawing.Point(1519, 317)
        Me.KCBL_GF.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_GF.Name = "KCBL_GF"
        Me.KCBL_GF.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_GF.TabIndex = 178
        Me.KCBL_GF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label31
        '
        Me.Label31.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label31.Location = New System.Drawing.Point(1353, 307)
        Me.Label31.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label31.Name = "Label31"
        Me.Label31.Size = New System.Drawing.Size(154, 56)
        Me.Label31.TabIndex = 177
        Me.Label31.Text = "扣除比例(%)"
        Me.Label31.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_GF
        '
        Me.YYNX_GF.Location = New System.Drawing.Point(1190, 317)
        Me.YYNX_GF.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_GF.Name = "YYNX_GF"
        Me.YYNX_GF.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_GF.TabIndex = 176
        Me.YYNX_GF.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label32
        '
        Me.Label32.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label32.Location = New System.Drawing.Point(1029, 308)
        Me.Label32.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label32.Name = "Label32"
        Me.Label32.Size = New System.Drawing.Size(158, 56)
        Me.Label32.TabIndex = 175
        Me.Label32.Text = "运营年限(年)"
        Me.Label32.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_NT
        '
        Me.KCBL_NT.Location = New System.Drawing.Point(1519, 261)
        Me.KCBL_NT.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_NT.Name = "KCBL_NT"
        Me.KCBL_NT.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_NT.TabIndex = 174
        Me.KCBL_NT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label29
        '
        Me.Label29.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label29.Location = New System.Drawing.Point(1353, 251)
        Me.Label29.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label29.Name = "Label29"
        Me.Label29.Size = New System.Drawing.Size(154, 56)
        Me.Label29.TabIndex = 173
        Me.Label29.Text = "扣除比例(%)"
        Me.Label29.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_NT
        '
        Me.YYNX_NT.Location = New System.Drawing.Point(1190, 261)
        Me.YYNX_NT.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_NT.Name = "YYNX_NT"
        Me.YYNX_NT.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_NT.TabIndex = 172
        Me.YYNX_NT.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label30
        '
        Me.Label30.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label30.Location = New System.Drawing.Point(1029, 252)
        Me.Label30.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label30.Name = "Label30"
        Me.Label30.Size = New System.Drawing.Size(158, 56)
        Me.Label30.TabIndex = 171
        Me.Label30.Text = "运营年限(年)"
        Me.Label30.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_RM
        '
        Me.KCBL_RM.Location = New System.Drawing.Point(1519, 205)
        Me.KCBL_RM.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_RM.Name = "KCBL_RM"
        Me.KCBL_RM.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_RM.TabIndex = 170
        Me.KCBL_RM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label24
        '
        Me.Label24.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label24.Location = New System.Drawing.Point(1353, 195)
        Me.Label24.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label24.Name = "Label24"
        Me.Label24.Size = New System.Drawing.Size(154, 56)
        Me.Label24.TabIndex = 169
        Me.Label24.Text = "扣除比例(%)"
        Me.Label24.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_RM
        '
        Me.YYNX_RM.Location = New System.Drawing.Point(1190, 205)
        Me.YYNX_RM.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_RM.Name = "YYNX_RM"
        Me.YYNX_RM.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_RM.TabIndex = 168
        Me.YYNX_RM.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label27
        '
        Me.Label27.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label27.Location = New System.Drawing.Point(1029, 196)
        Me.Label27.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label27.Name = "Label27"
        Me.Label27.Size = New System.Drawing.Size(158, 56)
        Me.Label27.TabIndex = 167
        Me.Label27.Text = "运营年限(年)"
        Me.Label27.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'KCBL_RJ
        '
        Me.KCBL_RJ.Location = New System.Drawing.Point(1519, 149)
        Me.KCBL_RJ.Margin = New System.Windows.Forms.Padding(6)
        Me.KCBL_RJ.Name = "KCBL_RJ"
        Me.KCBL_RJ.Size = New System.Drawing.Size(142, 35)
        Me.KCBL_RJ.TabIndex = 166
        Me.KCBL_RJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label28
        '
        Me.Label28.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label28.Location = New System.Drawing.Point(1353, 139)
        Me.Label28.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label28.Name = "Label28"
        Me.Label28.Size = New System.Drawing.Size(154, 56)
        Me.Label28.TabIndex = 165
        Me.Label28.Text = "扣除比例(%)"
        Me.Label28.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'YYNX_RJ
        '
        Me.YYNX_RJ.Location = New System.Drawing.Point(1190, 149)
        Me.YYNX_RJ.Margin = New System.Windows.Forms.Padding(6)
        Me.YYNX_RJ.Name = "YYNX_RJ"
        Me.YYNX_RJ.Size = New System.Drawing.Size(142, 35)
        Me.YYNX_RJ.TabIndex = 164
        Me.YYNX_RJ.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label26
        '
        Me.Label26.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label26.Location = New System.Drawing.Point(1029, 140)
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
        Me.fengdian.Location = New System.Drawing.Point(881, 375)
        Me.fengdian.Name = "fengdian"
        Me.fengdian.Size = New System.Drawing.Size(90, 28)
        Me.fengdian.TabIndex = 161
        Me.fengdian.Text = "风电"
        Me.fengdian.UseVisualStyleBackColor = True
        '
        'guangfu
        '
        Me.guangfu.AutoSize = True
        Me.guangfu.Location = New System.Drawing.Point(881, 323)
        Me.guangfu.Name = "guangfu"
        Me.guangfu.Size = New System.Drawing.Size(90, 28)
        Me.guangfu.TabIndex = 160
        Me.guangfu.Text = "光伏"
        Me.guangfu.UseVisualStyleBackColor = True
        '
        'nuantong
        '
        Me.nuantong.AutoSize = True
        Me.nuantong.Location = New System.Drawing.Point(881, 267)
        Me.nuantong.Name = "nuantong"
        Me.nuantong.Size = New System.Drawing.Size(90, 28)
        Me.nuantong.TabIndex = 159
        Me.nuantong.Text = "暖通"
        Me.nuantong.UseVisualStyleBackColor = True
        '
        'ranmei
        '
        Me.ranmei.AutoSize = True
        Me.ranmei.Location = New System.Drawing.Point(881, 209)
        Me.ranmei.Name = "ranmei"
        Me.ranmei.Size = New System.Drawing.Size(90, 28)
        Me.ranmei.TabIndex = 158
        Me.ranmei.Text = "燃煤"
        Me.ranmei.UseVisualStyleBackColor = True
        '
        'ranji
        '
        Me.ranji.AutoSize = True
        Me.ranji.Location = New System.Drawing.Point(881, 155)
        Me.ranji.Name = "ranji"
        Me.ranji.Size = New System.Drawing.Size(90, 28)
        Me.ranji.TabIndex = 157
        Me.ranji.Text = "燃机"
        Me.ranji.UseVisualStyleBackColor = True
        '
        'jsfl5
        '
        Me.jsfl5.Location = New System.Drawing.Point(643, 580)
        Me.jsfl5.Margin = New System.Windows.Forms.Padding(6)
        Me.jsfl5.Name = "jsfl5"
        Me.jsfl5.Size = New System.Drawing.Size(142, 35)
        Me.jsfl5.TabIndex = 156
        Me.jsfl5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksfl5
        '
        Me.ksfl5.Location = New System.Drawing.Point(419, 580)
        Me.ksfl5.Margin = New System.Windows.Forms.Padding(6)
        Me.ksfl5.Name = "ksfl5"
        Me.ksfl5.Size = New System.Drawing.Size(142, 35)
        Me.ksfl5.TabIndex = 155
        Me.ksfl5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsnf5
        '
        Me.jsnf5.Location = New System.Drawing.Point(213, 580)
        Me.jsnf5.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf5.Name = "jsnf5"
        Me.jsnf5.Size = New System.Drawing.Size(142, 35)
        Me.jsnf5.TabIndex = 152
        Me.jsnf5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label22
        '
        Me.Label22.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label22.Location = New System.Drawing.Point(225, 522)
        Me.Label22.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label22.Name = "Label22"
        Me.Label22.Size = New System.Drawing.Size(132, 56)
        Me.Label22.TabIndex = 151
        Me.Label22.Text = "结束年份"
        Me.Label22.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label23
        '
        Me.Label23.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label23.Location = New System.Drawing.Point(51, 518)
        Me.Label23.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label23.Name = "Label23"
        Me.Label23.Size = New System.Drawing.Size(134, 56)
        Me.Label23.TabIndex = 150
        Me.Label23.Text = "开始年份"
        Me.Label23.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ksnf5
        '
        Me.ksnf5.Location = New System.Drawing.Point(39, 580)
        Me.ksnf5.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf5.Name = "ksnf5"
        Me.ksnf5.Size = New System.Drawing.Size(142, 35)
        Me.ksnf5.TabIndex = 149
        Me.ksnf5.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '设置剔除
        '
        Me.设置剔除.BackColor = System.Drawing.SystemColors.ControlLight
        Me.设置剔除.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.设置剔除.Location = New System.Drawing.Point(881, 492)
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
        Me.清空窗体.Location = New System.Drawing.Point(1157, 492)
        Me.清空窗体.Margin = New System.Windows.Forms.Padding(6)
        Me.清空窗体.Name = "清空窗体"
        Me.清空窗体.Size = New System.Drawing.Size(224, 82)
        Me.清空窗体.TabIndex = 147
        Me.清空窗体.Text = "清空窗体"
        Me.清空窗体.UseVisualStyleBackColor = False
        '
        'RichTextBox1
        '
        Me.RichTextBox1.Location = New System.Drawing.Point(881, 667)
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
        Me.Label19.Location = New System.Drawing.Point(980, 599)
        Me.Label19.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label19.Name = "Label19"
        Me.Label19.Size = New System.Drawing.Size(583, 50)
        Me.Label19.TabIndex = 145
        Me.Label19.Text = "逐年材料费率OR其它费率计算结果显示"
        Me.Label19.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'jsfl4
        '
        Me.jsfl4.Location = New System.Drawing.Point(643, 481)
        Me.jsfl4.Margin = New System.Windows.Forms.Padding(6)
        Me.jsfl4.Name = "jsfl4"
        Me.jsfl4.Size = New System.Drawing.Size(142, 35)
        Me.jsfl4.TabIndex = 144
        Me.jsfl4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksfl4
        '
        Me.ksfl4.Location = New System.Drawing.Point(419, 481)
        Me.ksfl4.Margin = New System.Windows.Forms.Padding(6)
        Me.ksfl4.Name = "ksfl4"
        Me.ksfl4.Size = New System.Drawing.Size(142, 35)
        Me.ksfl4.TabIndex = 143
        Me.ksfl4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsnf4
        '
        Me.jsnf4.Location = New System.Drawing.Point(213, 481)
        Me.jsnf4.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf4.Name = "jsnf4"
        Me.jsnf4.Size = New System.Drawing.Size(142, 35)
        Me.jsnf4.TabIndex = 140
        Me.jsnf4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label17
        '
        Me.Label17.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label17.Location = New System.Drawing.Point(225, 423)
        Me.Label17.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label17.Name = "Label17"
        Me.Label17.Size = New System.Drawing.Size(132, 56)
        Me.Label17.TabIndex = 139
        Me.Label17.Text = "结束年份"
        Me.Label17.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label18
        '
        Me.Label18.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label18.Location = New System.Drawing.Point(51, 419)
        Me.Label18.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label18.Name = "Label18"
        Me.Label18.Size = New System.Drawing.Size(134, 56)
        Me.Label18.TabIndex = 138
        Me.Label18.Text = "开始年份"
        Me.Label18.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ksnf4
        '
        Me.ksnf4.Location = New System.Drawing.Point(39, 481)
        Me.ksnf4.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf4.Name = "ksnf4"
        Me.ksnf4.Size = New System.Drawing.Size(142, 35)
        Me.ksnf4.TabIndex = 137
        Me.ksnf4.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsfl3
        '
        Me.jsfl3.Location = New System.Drawing.Point(643, 380)
        Me.jsfl3.Margin = New System.Windows.Forms.Padding(6)
        Me.jsfl3.Name = "jsfl3"
        Me.jsfl3.Size = New System.Drawing.Size(142, 35)
        Me.jsfl3.TabIndex = 136
        Me.jsfl3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksfl3
        '
        Me.ksfl3.Location = New System.Drawing.Point(419, 380)
        Me.ksfl3.Margin = New System.Windows.Forms.Padding(6)
        Me.ksfl3.Name = "ksfl3"
        Me.ksfl3.Size = New System.Drawing.Size(142, 35)
        Me.ksfl3.TabIndex = 135
        Me.ksfl3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsnf3
        '
        Me.jsnf3.Location = New System.Drawing.Point(213, 380)
        Me.jsnf3.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf3.Name = "jsnf3"
        Me.jsnf3.Size = New System.Drawing.Size(142, 35)
        Me.jsnf3.TabIndex = 132
        Me.jsnf3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label13
        '
        Me.Label13.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label13.Location = New System.Drawing.Point(225, 322)
        Me.Label13.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(132, 56)
        Me.Label13.TabIndex = 131
        Me.Label13.Text = "结束年份"
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label14
        '
        Me.Label14.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label14.Location = New System.Drawing.Point(51, 318)
        Me.Label14.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label14.Name = "Label14"
        Me.Label14.Size = New System.Drawing.Size(134, 56)
        Me.Label14.TabIndex = 130
        Me.Label14.Text = "开始年份"
        Me.Label14.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ksnf3
        '
        Me.ksnf3.Location = New System.Drawing.Point(39, 380)
        Me.ksnf3.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf3.Name = "ksnf3"
        Me.ksnf3.Size = New System.Drawing.Size(142, 35)
        Me.ksnf3.TabIndex = 129
        Me.ksnf3.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsfl2
        '
        Me.jsfl2.Location = New System.Drawing.Point(643, 281)
        Me.jsfl2.Margin = New System.Windows.Forms.Padding(6)
        Me.jsfl2.Name = "jsfl2"
        Me.jsfl2.Size = New System.Drawing.Size(142, 35)
        Me.jsfl2.TabIndex = 128
        Me.jsfl2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksfl2
        '
        Me.ksfl2.Location = New System.Drawing.Point(419, 281)
        Me.ksfl2.Margin = New System.Windows.Forms.Padding(6)
        Me.ksfl2.Name = "ksfl2"
        Me.ksfl2.Size = New System.Drawing.Size(142, 35)
        Me.ksfl2.TabIndex = 127
        Me.ksfl2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsnf2
        '
        Me.jsnf2.Location = New System.Drawing.Point(213, 281)
        Me.jsnf2.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf2.Name = "jsnf2"
        Me.jsnf2.Size = New System.Drawing.Size(142, 35)
        Me.jsnf2.TabIndex = 124
        Me.jsnf2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label9
        '
        Me.Label9.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label9.Location = New System.Drawing.Point(225, 226)
        Me.Label9.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label9.Name = "Label9"
        Me.Label9.Size = New System.Drawing.Size(132, 56)
        Me.Label9.TabIndex = 123
        Me.Label9.Text = "结束年份"
        Me.Label9.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label10
        '
        Me.Label10.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label10.Location = New System.Drawing.Point(51, 222)
        Me.Label10.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label10.Name = "Label10"
        Me.Label10.Size = New System.Drawing.Size(134, 56)
        Me.Label10.TabIndex = 122
        Me.Label10.Text = "开始年份"
        Me.Label10.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ksnf2
        '
        Me.ksnf2.Location = New System.Drawing.Point(39, 281)
        Me.ksnf2.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf2.Name = "ksnf2"
        Me.ksnf2.Size = New System.Drawing.Size(142, 35)
        Me.ksnf2.TabIndex = 121
        Me.ksnf2.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'jsfl1
        '
        Me.jsfl1.Location = New System.Drawing.Point(643, 183)
        Me.jsfl1.Margin = New System.Windows.Forms.Padding(6)
        Me.jsfl1.Name = "jsfl1"
        Me.jsfl1.Size = New System.Drawing.Size(142, 35)
        Me.jsfl1.TabIndex = 120
        Me.jsfl1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'ksfl1
        '
        Me.ksfl1.Location = New System.Drawing.Point(419, 183)
        Me.ksfl1.Margin = New System.Windows.Forms.Padding(6)
        Me.ksfl1.Name = "ksfl1"
        Me.ksfl1.Size = New System.Drawing.Size(142, 35)
        Me.ksfl1.TabIndex = 119
        Me.ksfl1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label2
        '
        Me.Label2.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label2.Location = New System.Drawing.Point(622, 125)
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
        Me.Label1.Location = New System.Drawing.Point(401, 125)
        Me.Label1.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(178, 56)
        Me.Label1.TabIndex = 117
        Me.Label1.Text = "开始年份费率"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'jsnf1
        '
        Me.jsnf1.Location = New System.Drawing.Point(213, 183)
        Me.jsnf1.Margin = New System.Windows.Forms.Padding(6)
        Me.jsnf1.Name = "jsnf1"
        Me.jsnf1.Size = New System.Drawing.Size(142, 35)
        Me.jsnf1.TabIndex = 116
        Me.jsnf1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label4
        '
        Me.Label4.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label4.Location = New System.Drawing.Point(225, 121)
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
        Me.Label3.Location = New System.Drawing.Point(51, 124)
        Me.Label3.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(134, 56)
        Me.Label3.TabIndex = 114
        Me.Label3.Text = "开始年份"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ksnf1
        '
        Me.ksnf1.Location = New System.Drawing.Point(39, 183)
        Me.ksnf1.Margin = New System.Windows.Forms.Padding(6)
        Me.ksnf1.Name = "ksnf1"
        Me.ksnf1.Size = New System.Drawing.Size(142, 35)
        Me.ksnf1.TabIndex = 113
        Me.ksnf1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'Label7
        '
        Me.Label7.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label7.Location = New System.Drawing.Point(189, 75)
        Me.Label7.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label7.Name = "Label7"
        Me.Label7.Size = New System.Drawing.Size(528, 50)
        Me.Label7.TabIndex = 112
        Me.Label7.Text = "输入开始和结束变化的年份以及逐年变化率"
        Me.Label7.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label5
        '
        Me.Label5.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label5.Location = New System.Drawing.Point(622, 222)
        Me.Label5.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label5.Name = "Label5"
        Me.Label5.Size = New System.Drawing.Size(196, 56)
        Me.Label5.TabIndex = 200
        Me.Label5.Text = "结束年份费率"
        Me.Label5.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label6
        '
        Me.Label6.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label6.Location = New System.Drawing.Point(401, 222)
        Me.Label6.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label6.Name = "Label6"
        Me.Label6.Size = New System.Drawing.Size(178, 56)
        Me.Label6.TabIndex = 199
        Me.Label6.Text = "开始年份费率"
        Me.Label6.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label11
        '
        Me.Label11.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label11.Location = New System.Drawing.Point(622, 323)
        Me.Label11.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label11.Name = "Label11"
        Me.Label11.Size = New System.Drawing.Size(196, 56)
        Me.Label11.TabIndex = 202
        Me.Label11.Text = "结束年份费率"
        Me.Label11.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label12
        '
        Me.Label12.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label12.Location = New System.Drawing.Point(401, 323)
        Me.Label12.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(178, 56)
        Me.Label12.TabIndex = 201
        Me.Label12.Text = "开始年份费率"
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label15
        '
        Me.Label15.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label15.Location = New System.Drawing.Point(622, 421)
        Me.Label15.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label15.Name = "Label15"
        Me.Label15.Size = New System.Drawing.Size(196, 56)
        Me.Label15.TabIndex = 204
        Me.Label15.Text = "结束年份费率"
        Me.Label15.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label16
        '
        Me.Label16.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label16.Location = New System.Drawing.Point(401, 421)
        Me.Label16.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label16.Name = "Label16"
        Me.Label16.Size = New System.Drawing.Size(178, 56)
        Me.Label16.TabIndex = 203
        Me.Label16.Text = "开始年份费率"
        Me.Label16.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label20
        '
        Me.Label20.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label20.Location = New System.Drawing.Point(622, 522)
        Me.Label20.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label20.Name = "Label20"
        Me.Label20.Size = New System.Drawing.Size(196, 56)
        Me.Label20.TabIndex = 206
        Me.Label20.Text = "结束年份费率"
        Me.Label20.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label21
        '
        Me.Label21.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label21.Location = New System.Drawing.Point(401, 522)
        Me.Label21.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label21.Name = "Label21"
        Me.Label21.Size = New System.Drawing.Size(178, 56)
        Me.Label21.TabIndex = 205
        Me.Label21.Text = "开始年份费率"
        Me.Label21.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'ComboBox1
        '
        Me.ComboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.ComboBox1.FormattingEnabled = True
        Me.ComboBox1.Items.AddRange(New Object() {"材料费率", "其它费率"})
        Me.ComboBox1.Location = New System.Drawing.Point(542, 644)
        Me.ComboBox1.Name = "ComboBox1"
        Me.ComboBox1.Size = New System.Drawing.Size(243, 32)
        Me.ComboBox1.TabIndex = 208
        '
        'Label38
        '
        Me.Label38.Font = New System.Drawing.Font("宋体", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Label38.Location = New System.Drawing.Point(35, 634)
        Me.Label38.Margin = New System.Windows.Forms.Padding(6, 0, 6, 0)
        Me.Label38.Name = "Label38"
        Me.Label38.Size = New System.Drawing.Size(445, 50)
        Me.Label38.TabIndex = 207
        Me.Label38.Text = "选择需要设置的费率类型"
        Me.Label38.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        '设置材料费和其它费计算方式
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(1698, 930)
        Me.Controls.Add(Me.ComboBox1)
        Me.Controls.Add(Me.Label38)
        Me.Controls.Add(Me.Label20)
        Me.Controls.Add(Me.Label21)
        Me.Controls.Add(Me.Label15)
        Me.Controls.Add(Me.Label16)
        Me.Controls.Add(Me.Label11)
        Me.Controls.Add(Me.Label12)
        Me.Controls.Add(Me.Label5)
        Me.Controls.Add(Me.Label6)
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
        Me.Controls.Add(Me.jsfl5)
        Me.Controls.Add(Me.ksfl5)
        Me.Controls.Add(Me.jsnf5)
        Me.Controls.Add(Me.Label22)
        Me.Controls.Add(Me.Label23)
        Me.Controls.Add(Me.ksnf5)
        Me.Controls.Add(Me.设置剔除)
        Me.Controls.Add(Me.清空窗体)
        Me.Controls.Add(Me.RichTextBox1)
        Me.Controls.Add(Me.Label19)
        Me.Controls.Add(Me.jsfl4)
        Me.Controls.Add(Me.ksfl4)
        Me.Controls.Add(Me.jsnf4)
        Me.Controls.Add(Me.Label17)
        Me.Controls.Add(Me.Label18)
        Me.Controls.Add(Me.ksnf4)
        Me.Controls.Add(Me.jsfl3)
        Me.Controls.Add(Me.ksfl3)
        Me.Controls.Add(Me.jsnf3)
        Me.Controls.Add(Me.Label13)
        Me.Controls.Add(Me.Label14)
        Me.Controls.Add(Me.ksnf3)
        Me.Controls.Add(Me.jsfl2)
        Me.Controls.Add(Me.ksfl2)
        Me.Controls.Add(Me.jsnf2)
        Me.Controls.Add(Me.Label9)
        Me.Controls.Add(Me.Label10)
        Me.Controls.Add(Me.ksnf2)
        Me.Controls.Add(Me.jsfl1)
        Me.Controls.Add(Me.ksfl1)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.jsnf1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.ksnf1)
        Me.Controls.Add(Me.Label7)
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
    Friend WithEvents jsfl5 As System.Windows.Forms.TextBox
    Friend WithEvents ksfl5 As System.Windows.Forms.TextBox
    Friend WithEvents jsnf5 As System.Windows.Forms.TextBox
    Friend WithEvents Label22 As System.Windows.Forms.Label
    Friend WithEvents Label23 As System.Windows.Forms.Label
    Friend WithEvents ksnf5 As System.Windows.Forms.TextBox
    Friend WithEvents 设置剔除 As System.Windows.Forms.Button
    Friend WithEvents 清空窗体 As System.Windows.Forms.Button
    Friend WithEvents RichTextBox1 As System.Windows.Forms.RichTextBox
    Friend WithEvents Label19 As System.Windows.Forms.Label
    Friend WithEvents jsfl4 As System.Windows.Forms.TextBox
    Friend WithEvents ksfl4 As System.Windows.Forms.TextBox
    Friend WithEvents jsnf4 As System.Windows.Forms.TextBox
    Friend WithEvents Label17 As System.Windows.Forms.Label
    Friend WithEvents Label18 As System.Windows.Forms.Label
    Friend WithEvents ksnf4 As System.Windows.Forms.TextBox
    Friend WithEvents jsfl3 As System.Windows.Forms.TextBox
    Friend WithEvents ksfl3 As System.Windows.Forms.TextBox
    Friend WithEvents jsnf3 As System.Windows.Forms.TextBox
    Friend WithEvents Label13 As System.Windows.Forms.Label
    Friend WithEvents Label14 As System.Windows.Forms.Label
    Friend WithEvents ksnf3 As System.Windows.Forms.TextBox
    Friend WithEvents jsfl2 As System.Windows.Forms.TextBox
    Friend WithEvents ksfl2 As System.Windows.Forms.TextBox
    Friend WithEvents jsnf2 As System.Windows.Forms.TextBox
    Friend WithEvents Label9 As System.Windows.Forms.Label
    Friend WithEvents Label10 As System.Windows.Forms.Label
    Friend WithEvents ksnf2 As System.Windows.Forms.TextBox
    Friend WithEvents jsfl1 As System.Windows.Forms.TextBox
    Friend WithEvents ksfl1 As System.Windows.Forms.TextBox
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents jsnf1 As System.Windows.Forms.TextBox
    Friend WithEvents Label4 As System.Windows.Forms.Label
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents ksnf1 As System.Windows.Forms.TextBox
    Friend WithEvents Label7 As System.Windows.Forms.Label
    Friend WithEvents Label5 As System.Windows.Forms.Label
    Friend WithEvents Label6 As System.Windows.Forms.Label
    Friend WithEvents Label11 As System.Windows.Forms.Label
    Friend WithEvents Label12 As System.Windows.Forms.Label
    Friend WithEvents Label15 As System.Windows.Forms.Label
    Friend WithEvents Label16 As System.Windows.Forms.Label
    Friend WithEvents Label20 As System.Windows.Forms.Label
    Friend WithEvents Label21 As System.Windows.Forms.Label
    Friend WithEvents ComboBox1 As System.Windows.Forms.ComboBox
    Friend WithEvents Label38 As System.Windows.Forms.Label
End Class
