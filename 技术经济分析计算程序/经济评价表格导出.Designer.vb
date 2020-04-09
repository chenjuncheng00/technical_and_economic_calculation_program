<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 经济评价表格导出
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
        Me.Checkbox1 = New System.Windows.Forms.CheckBox()
        Me.CheckBox2 = New System.Windows.Forms.CheckBox()
        Me.CheckBox3 = New System.Windows.Forms.CheckBox()
        Me.CheckBox4 = New System.Windows.Forms.CheckBox()
        Me.确定 = New System.Windows.Forms.Button()
        Me.CheckBox5 = New System.Windows.Forms.CheckBox()
        Me.SuspendLayout()
        '
        'Checkbox1
        '
        Me.Checkbox1.AutoSize = True
        Me.Checkbox1.Checked = True
        Me.Checkbox1.CheckState = System.Windows.Forms.CheckState.Checked
        Me.Checkbox1.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.Checkbox1.Location = New System.Drawing.Point(95, 58)
        Me.Checkbox1.Margin = New System.Windows.Forms.Padding(6)
        Me.Checkbox1.Name = "Checkbox1"
        Me.Checkbox1.Size = New System.Drawing.Size(408, 32)
        Me.Checkbox1.TabIndex = 136
        Me.Checkbox1.Text = "清除全部表格中的单元格批注"
        Me.Checkbox1.UseVisualStyleBackColor = True
        '
        'CheckBox2
        '
        Me.CheckBox2.AutoSize = True
        Me.CheckBox2.Checked = True
        Me.CheckBox2.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox2.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox2.Location = New System.Drawing.Point(95, 132)
        Me.CheckBox2.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox2.Name = "CheckBox2"
        Me.CheckBox2.Size = New System.Drawing.Size(380, 32)
        Me.CheckBox2.TabIndex = 137
        Me.CheckBox2.Text = "导出全部经济评价详细表格"
        Me.CheckBox2.UseVisualStyleBackColor = True
        '
        'CheckBox3
        '
        Me.CheckBox3.AutoSize = True
        Me.CheckBox3.Checked = True
        Me.CheckBox3.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox3.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox3.Location = New System.Drawing.Point(95, 278)
        Me.CheckBox3.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox3.Name = "CheckBox3"
        Me.CheckBox3.Size = New System.Drawing.Size(464, 32)
        Me.CheckBox3.TabIndex = 138
        Me.CheckBox3.Text = "导出敏感性分析表、敏感性分析图"
        Me.CheckBox3.UseVisualStyleBackColor = True
        '
        'CheckBox4
        '
        Me.CheckBox4.AutoSize = True
        Me.CheckBox4.Checked = True
        Me.CheckBox4.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox4.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox4.Location = New System.Drawing.Point(95, 353)
        Me.CheckBox4.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox4.Name = "CheckBox4"
        Me.CheckBox4.Size = New System.Drawing.Size(520, 32)
        Me.CheckBox4.TabIndex = 139
        Me.CheckBox4.Text = "导出盈亏平衡分析表、盈亏平衡分析图"
        Me.CheckBox4.UseVisualStyleBackColor = True
        '
        '确定
        '
        Me.确定.BackColor = System.Drawing.SystemColors.ControlLight
        Me.确定.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.确定.Location = New System.Drawing.Point(249, 426)
        Me.确定.Margin = New System.Windows.Forms.Padding(6)
        Me.确定.Name = "确定"
        Me.确定.Size = New System.Drawing.Size(224, 82)
        Me.确定.TabIndex = 140
        Me.确定.Text = "确定"
        Me.确定.UseVisualStyleBackColor = False
        '
        'CheckBox5
        '
        Me.CheckBox5.AutoSize = True
        Me.CheckBox5.Checked = True
        Me.CheckBox5.CheckState = System.Windows.Forms.CheckState.Checked
        Me.CheckBox5.Font = New System.Drawing.Font("宋体", 10.5!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.CheckBox5.Location = New System.Drawing.Point(95, 203)
        Me.CheckBox5.Margin = New System.Windows.Forms.Padding(6)
        Me.CheckBox5.Name = "CheckBox5"
        Me.CheckBox5.Size = New System.Drawing.Size(352, 32)
        Me.CheckBox5.TabIndex = 141
        Me.CheckBox5.Text = "导出财务评价指标一览表"
        Me.CheckBox5.UseVisualStyleBackColor = True
        '
        '经济评价表格导出
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(12.0!, 24.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(714, 548)
        Me.Controls.Add(Me.CheckBox5)
        Me.Controls.Add(Me.确定)
        Me.Controls.Add(Me.CheckBox4)
        Me.Controls.Add(Me.CheckBox3)
        Me.Controls.Add(Me.CheckBox2)
        Me.Controls.Add(Me.Checkbox1)
        Me.Name = "经济评价表格导出"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "经济评价表格导出"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Checkbox1 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox2 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox3 As System.Windows.Forms.CheckBox
    Friend WithEvents CheckBox4 As System.Windows.Forms.CheckBox
    Friend WithEvents 确定 As System.Windows.Forms.Button
    Friend WithEvents CheckBox5 As System.Windows.Forms.CheckBox
End Class
