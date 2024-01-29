Imports Microsoft.Office.Interop.Excel
Public Class 设置流动资金计算方式
    Private Sub 设置流动资金计算方式_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        'Checkbox载入成本文字
        Me.cb_rldl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 9).Value, String)
        Me.cb_rldl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 9).Value, String)
        Me.cb_rldl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 9).Value, String)
        Me.cb_rldl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 9).Value, String)
        Me.cb_rldl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 9).Value, String)
        Me.cb_ycl1.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 9).Value, String)
        Me.cb_ycl2.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 9).Value, String)
        Me.cb_ycl3.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 9).Value, String)
        Me.cb_ycl4.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 9).Value, String)
        Me.cb_ycl5.Text = CType(ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 9).Value, String)
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '根据输入的成本等内容，判断checkbox是否可以被勾选     
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 14).Value > 0 Then
            Me.cb_rldl1.Enabled = True
        Else
            Me.cb_rldl1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 14).Value > 0 Then
            Me.cb_rldl2.Enabled = True
        Else
            Me.cb_rldl2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 14).Value > 0 Then
            Me.cb_rldl3.Enabled = True
        Else
            Me.cb_rldl3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 14).Value > 0 Then
            Me.cb_rldl4.Enabled = True
        Else
            Me.cb_rldl4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 14).Value > 0 Then
            Me.cb_rldl5.Enabled = True
        Else
            Me.cb_rldl5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 14).Value > 0 Then
            Me.cb_ycl1.Enabled = True
        Else
            Me.cb_ycl1.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 14).Value > 0 Then
            Me.cb_ycl2.Enabled = True
        Else
            Me.cb_ycl2.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 14).Value > 0 Then
            Me.cb_ycl3.Enabled = True
        Else
            Me.cb_ycl3.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 14).Value > 0 Then
            Me.cb_ycl4.Enabled = True
        Else
            Me.cb_ycl4.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 14).Value > 0 Then
            Me.cb_ycl5.Enabled = True
        Else
            Me.cb_ycl5.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 14).Value > 0 Then
            Me.cb_rygz.Enabled = True
            Me.rgf_ft.Enabled = True
        Else
            Me.cb_rygz.Enabled = False
            Me.rgf_ft.Enabled = False
        End If
        '根据输入的分项建设等内容，判断checkbox是否可以被勾选     
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(44, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(45, 1).Value > 0 Then
            Me.jsnr_rj.Enabled = True
            Me.YYNX_RJ.Enabled = True
            Me.KCBL_RJ.Enabled = True
        Else
            Me.jsnr_rj.Enabled = False
            Me.YYNX_RJ.Enabled = False
            Me.KCBL_RJ.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(46, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(47, 1).Value > 0 Then
            Me.jsnr_xdc.Enabled = True
            Me.YYNX_XDC.Enabled = True
            Me.KCBL_XDC.Enabled = True
        Else
            Me.jsnr_xdc.Enabled = False
            Me.YYNX_XDC.Enabled = False
            Me.KCBL_XDC.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(48, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(49, 1).Value > 0 Then
            Me.jsnr_nt.Enabled = True
            Me.YYNX_NT.Enabled = True
            Me.KCBL_NT.Enabled = True
        Else
            Me.jsnr_nt.Enabled = False
            Me.YYNX_NT.Enabled = False
            Me.KCBL_NT.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(50, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(51, 1).Value > 0 Then
            Me.jsnr_gf.Enabled = True
            Me.YYNX_GF.Enabled = True
            Me.KCBL_GF.Enabled = True
        Else
            Me.jsnr_gf.Enabled = False
            Me.YYNX_GF.Enabled = False
            Me.KCBL_GF.Enabled = False
        End If
        If ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(53, 1).Value + ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(54, 1).Value > 0 Then
            Me.jsnr_fd.Enabled = True
            Me.YYNX_FD.Enabled = True
            Me.KCBL_FD.Enabled = True
        Else
            Me.jsnr_fd.Enabled = False
            Me.YYNX_FD.Enabled = False
            Me.KCBL_FD.Enabled = False
        End If
    End Sub

    Private Sub 重置默认_Click(sender As Object, e As EventArgs) Handles 重置默认.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————————————————————————————————
        '————————————————————————————————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否将流动资金计算方式重置回默认设置？？", vbOKCancel)
        If XZ = vbOK Then
            '项目计算年限
            Dim jsnx As Integer = ExcelApp.ThisWorkbook.Worksheets("估算表").Cells(5, 7).Value
            '计算基数扣除默认设置
            Dim kcje_mr_ldzj = 流动资金计算基数扣除默认设置(jsnx)
            '运营年限
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 3).Value = kcje_mr_ldzj(0)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 4).Value = kcje_mr_ldzj(1)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 5).Value = kcje_mr_ldzj(2)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 6).Value = kcje_mr_ldzj(3)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 7).Value = kcje_mr_ldzj(4)
            '扣除比例
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 3).Value = kcje_mr_ldzj(5)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 4).Value = kcje_mr_ldzj(6)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 5).Value = kcje_mr_ldzj(7)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 6).Value = kcje_mr_ldzj(8)
            ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 7).Value = kcje_mr_ldzj(9)
            '流动资金常规计算模式设置，0：采用常规计算模型；1：采用新能源计算模式
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(30, 33).Value = 0
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(32, 33).Value = 0
            '显示计算模式
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(59, 18).Value = "常规设置"
            '———————————————————————————————————————————————————————————————————————————————————————— 
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '收入和成本变化后相关计算
            Call 收入成本相关计算(ExcelApp, sdsl_model, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '———————————————————————————————————————————————————————————————————————————————————————— 
            MsgBox("已经将流动资金计算方式重置回默认设置！")
        End If
    End Sub

    Private Sub 计算_Click(sender As Object, e As EventArgs) Handles 计算.Click
        On Error Resume Next
        '定义Excel对象
        Dim ExcelApp As Application '定义Excel对象
        ExcelApp = GetObject(, "Excel.Application")    '当前EXCEL对象赋值给ExcelApp
        '————————————————————————————————————————————————————————————————————
        Dim XZ = MsgBox("是否确认选择的各项内容？", vbOKCancel)
        If XZ = vbOK Then
            '成本内容设置
            If Me.cb_rldl1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(14, 33).Value = 0
            End If
            If Me.cb_rldl2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(15, 33).Value = 0
            End If
            If Me.cb_rldl3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(16, 33).Value = 0
            End If
            If Me.cb_rldl4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(17, 33).Value = 0
            End If
            If Me.cb_rldl5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(19, 33).Value = 0
            End If
            If Me.cb_ycl1.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(21, 33).Value = 0
            End If
            If Me.cb_ycl2.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(22, 33).Value = 0
            End If
            If Me.cb_ycl3.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(23, 33).Value = 0
            End If
            If Me.cb_ycl4.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(24, 33).Value = 0
            End If
            If Me.cb_ycl5.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 33).Value = 1
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(25, 33).Value = 0
            End If
            If Me.cb_rygz.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 33).Value = CType(Me.rgf_ft.Text, Double) / 100
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(27, 33).Value = 0
            End If
            '————————————————————————————————————————————————————————————————————
            '投资内容设置
            '燃机
            If Me.jsnr_rj.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(30, 33).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 3).Value = CType(Me.YYNX_RJ.Text, Integer)
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 3).Value = CType(Me.KCBL_RJ.Text, Double) / 100
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(30, 33).Value = 0
            End If
            '蓄电池
            If Me.jsnr_xdc.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 4).Value = CType(Me.YYNX_XDC.Text, Integer)
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 4).Value = CType(Me.KCBL_XDC.Text, Double) / 100
            End If
            '暖通
            If Me.jsnr_nt.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(32, 33).Value = 1
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 5).Value = CType(Me.YYNX_NT.Text, Integer)
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 5).Value = CType(Me.KCBL_NT.Text, Double) / 100
            Else
                ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(32, 33).Value = 0
            End If
            '光伏
            If Me.jsnr_gf.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 6).Value = CType(Me.YYNX_GF.Text, Integer)
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 6).Value = CType(Me.KCBL_GF.Text, Double) / 100
            End If
            '风电
            If Me.jsnr_fd.Checked = True Then
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(56, 7).Value = CType(Me.YYNX_FD.Text, Integer)
                ExcelApp.ThisWorkbook.Worksheets("流动资金估算表").Cells(57, 7).Value = CType(Me.KCBL_FD.Text, Double) / 100
            End If
            '———————————————————————————————————————————————————————————————————————————————————————— 
            ExcelApp.ThisWorkbook.Worksheets("收入&成本输入").Cells(59, 18).Value = "特殊设置"
            '————————————————————————————————————————————————————————————————————
            '计算模式
            Dim ans_mode = 计算模式_从EXCEL读取(ExcelApp)
            Dim zbj_model = ans_mode(0)
            Dim hscz = ans_mode(1)
            Dim xlfl_cg_model = ans_mode(2)
            Dim xlfl_qt_model = ans_mode(3)
            Dim clfl_qtfl_model = ans_mode(4)
            Dim sdsl_model = ans_mode(5)
            Dim kcje_xlf_model = ans_mode(6)
            Dim kcje_clf_qtf_model = ans_mode(7)
            Dim ldzj_model = ans_mode(8)
            Dim kcje_ldzj_model = ans_mode(9)
            Dim bxf_model = ans_mode(10)
            Dim kcje_bxf_model = ans_mode(11)
            '————————————————————————————————————————————————————————————————————————————————————————
            '计算流动资金，包括：材料费其它费+修理费+保险费
            Call 流动资金相关计算(ExcelApp, xlfl_cg_model, xlfl_qt_model, kcje_xlf_model, hscz, zbj_model, clfl_qtfl_model, kcje_clf_qtf_model, ldzj_model, kcje_ldzj_model, bxf_model, kcje_bxf_model)
            '重新计算投资收益率和投资回收期
            Call 投资收益率和投资回收期计算(ExcelApp)
            MsgBox("计算完成！")
        End If
    End Sub
End Class