Imports Microsoft.Office.Interop

Public Class 进入维护模式
    Private Sub 确定_Click(sender As Object, e As EventArgs) Handles 确定.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Microsoft.Office.Interop.Excel.Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        Dim 技术经济分析计算程序 As New Com技术经济分析计算程序
        Dim mima As String '输入的密码
        mima = CType(Com技术经济分析计算程序.Form3.TextBox1.Text, String)
        If mima = "cjc19920105" Then
            '解锁工作表
            Call 技术经济分析计算程序.解锁表格()
            '显示工作表
            ExcelApp.Worksheets("投资方1现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            ExcelApp.Worksheets("投资方2现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            ExcelApp.Worksheets("投资方3现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            ExcelApp.Worksheets("投资方4现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            ExcelApp.Worksheets("投资方5现金流量表").Visible = Excel.XlSheetVisibility.xlSheetVisible
            '密码输入正确后，将计数器重置为0
            ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value = 0
            Com技术经济分析计算程序.Form3.Close()
            MsgBox("已成功进入维护模式，程序可以被编辑！")
        Else
            ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value = ExcelApp.Worksheets("建设期时间计划表").Cells(2, 26).Value + 1
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '提醒
            MsgBox("密码错误，请重新输入！")
            Com技术经济分析计算程序.Form3.Close()
            '锁定表格
            Call 技术经济分析计算程序.锁定表格()
            '保存表格的改动
            ExcelApp.Application.DisplayAlerts = False
            ExcelApp.ThisWorkbook.Save()
            ExcelApp.Application.DisplayAlerts = True
            '清空文本框中的输入内容
            Com技术经济分析计算程序.Form3.TextBox1.Clear()
        End If
        Me.Close()
    End Sub
    Protected Overrides Sub OnKeyPress(ByVal e As System.Windows.Forms.KeyPressEventArgs)
        Dim key As String
        key = e.KeyChar
        '检验按键是否为回车键，如果是就把焦点附给按钮1，并执行Click命令
        If key = Microsoft.VisualBasic.ChrW(13) Then
            确定.Focus()
            确定.PerformClick()
        End If
    End Sub
    Private Sub 进入维护模式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        MyBase.KeyPreview = True
    End Sub
End Class