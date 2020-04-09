<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class 进入维护模式
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
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.确定 = New System.Windows.Forms.Button()
        Me.SuspendLayout()
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(12, 12)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.TextBox1.Size = New System.Drawing.Size(160, 21)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        '确定
        '
        Me.确定.BackColor = System.Drawing.SystemColors.ControlLight
        Me.确定.Font = New System.Drawing.Font("宋体", 12.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(134, Byte))
        Me.确定.Location = New System.Drawing.Point(179, 3)
        Me.确定.Name = "确定"
        Me.确定.Size = New System.Drawing.Size(74, 37)
        Me.确定.TabIndex = 18
        Me.确定.Text = "确定"
        Me.确定.UseVisualStyleBackColor = False
        '
        '进入维护模式
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(259, 43)
        Me.Controls.Add(Me.确定)
        Me.Controls.Add(Me.TextBox1)
        Me.Name = "进入维护模式"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "进入维护模式"
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub
    Friend WithEvents TextBox1 As System.Windows.Forms.TextBox
    Friend WithEvents 确定 As System.Windows.Forms.Button
End Class
